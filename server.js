const express = require('express');
const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientCredentialsAuthProvider } = require('@microsoft/microsoft-graph-sdk/auth');
const mongoose = require('mongoose');
const bcrypt = require('bcrypt');
const jwt = require('jsonwebtoken');
const cors = require('cors');
const dotenv = require('dotenv');
const pdfParse = require('pdf-parse');
const mammoth = require('mammoth');

// Load environment variables
dotenv.config();

const app = express();
app.use(cors());
app.use(express.json());

// MongoDB Connection
mongoose.connect(process.env.MONGODB_URI, { useNewUrlParser: true, useUnifiedTopology: true })
  .then(() => console.log('Connected to MongoDB'))
  .catch(err => console.error('MongoDB connection error:', err));

// User Schema for Local Credentials
const userSchema = new mongoose.Schema({
  username: { type: String, unique: true, required: true },
  password: { type: String, required: true },
  role: { type: String, enum: ['user', 'admin'], default: 'user' }
});

const User = mongoose.model('User', userSchema);

// Microsoft Graph Client Setup
const client = Client.initWithMiddleware({
  authProvider: new ClientCredentialsAuthProvider({
    clientId: process.env.CLIENT_ID,
    clientSecret: process.env.CLIENT_SECRET,
    tenantId: process.env.TENANT_ID
  })
});

// Middleware to Verify JWT
const authenticateToken = (req, res, next) => {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.split(' ')[1];
  if (!token) return res.status(401).json({ error: 'Access denied' });

  jwt.verify(token, process.env.JWT_SECRET, (err, user) => {
    if (err) return res.status(403).json({ error: 'Invalid token' });
    req.user = user;
    next();
  });
};

// Middleware to Check Admin Role
const isAdmin = (req, res, next) => {
  if (req.user.role !== 'admin') return res.status(403).json({ error: 'Admin access required' });
  next();
};

// Login Endpoint
app.post('/api/login', async (req, res) => {
  const { username, password } = req.body;
  try {
    const user = await User.findOne({ username });
    if (!user) return res.status(400).json({ error: 'Invalid credentials' });

    const validPassword = await bcrypt.compare(password, user.password);
    if (!validPassword) return res.status(400).json({ error: 'Invalid credentials' });

    const token = jwt.sign({ username: user.username, role: user.role }, process.env.JWT_SECRET, { expiresIn: '1h' });
    res.json({ token });
  } catch (error) {
    res.status(500).json({ error: 'Server error' });
  }
});

// Admin: Create User
app.post('/api/admin/users', authenticateToken, isAdmin, async (req, res) => {
  const { username, password, role } = req.body;
  try {
    const hashedPassword = await bcrypt.hash(password, 10);
    const user = new User({ username, password: hashedPassword, role });
    await user.save();
    res.status(201).json({ message: 'User created' });
  } catch (error) {
    res.status(400).json({ error: 'Error creating user' });
  }
});

// Admin: Update User
app.put('/api/admin/users/:username', authenticateToken, isAdmin, async (req, res) => {
  const { password, role } = req.body;
  try {
    const updateData = {};
    if (password) updateData.password = await bcrypt.hash(password, 10);
    if (role) updateData.role = role;

    const user = await User.findOneAndUpdate({ username: req.params.username }, updateData, { new: true });
    if (!user) return res.status(404).json({ error: 'User not found' });
    res.json({ message: 'User updated' });
  } catch (error) {
    res.status(400).json({ error: 'Error updating user' });
  }
});

// Admin: Delete User
app.delete('/api/admin/users/:username', authenticateToken, isAdmin, async (req, res) => {
  try {
    const user = await User.findOneAndDelete({ username: req.params.username });
    if (!user) return res.status(404).json({ error: 'User not found' });
    res.json({ message: 'User deleted' });
  } catch (error) {
    res.status(400).json({ error: 'Error deleting user' });
  }
});

// Chatbot Query Endpoint
app.post('/api/chat', authenticateToken, async (req, res) => {
  const { query, practiceArea } = req.body;
  try {
    // Map practice areas to SharePoint folder paths
    const practiceAreaPaths = {
      'Corporate': '/sites/Adsero/Documents/Corporate',
      'DisputeResolution': '/sites/Adsero/Documents/DisputeResolution',
      'Tax': '/sites/Adsero/Documents/Tax',
      // Add other practice areas from adsero.me
    };

    const folderPath = practiceAreaPaths[practiceArea] || '/sites/Adsero/Documents/General';
    
    // Fetch files from SharePoint
    const files = await client.api(`/sites/adsero.sharepoint.com:/${folderPath}:/children`)
      .select('name,webUrl,file')
      .get();

    let documentsContent = [];
    for (const item of files.value) {
      if (item.file) {
        const fileContent = await client.api(`/sites/adsero.sharepoint.com:/${folderPath}/${item.name}:/content`).get();
        let text = '';

        if (item.name.endsWith('.pdf')) {
          const pdfData = await pdfParse(fileContent);
          text = pdfData.text;
        } else if (item.name.endsWith('.docx')) {
          const docData = await mammoth.extractRawText({ buffer: fileContent });
          text = docData.value;
        } else {
          text = fileContent.toString();
        }

        documentsContent.push({ name: item.name, content: text });
      }
    }

    // Simple keyword-based search (replace with AI embeddings for better results)
    const relevantDocs = documentsContent.filter(doc => doc.content.toLowerCase().includes(query.toLowerCase()));
    if (relevantDocs.length === 0) {
      return res.json({ response: 'No relevant information found.' });
    }

    // Combine content (limit to avoid token overflow)
    const combinedContent = relevantDocs.map(doc => doc.content).join('\n').substring(0, 4000);
    res.json({ response: combinedContent || 'Found documents but no specific answer extracted.' });
  } catch (error) {
    console.error('Chat error:', error);
    res.status(500).json({ error: 'Error processing query' });
  }
});

// Admin: Manage SharePoint Library (Example: List Libraries)
app.get('/api/admin/libraries', authenticateToken, isAdmin, async (req, res) => {
  try {
    const libraries = await client.api('/sites/adsero.sharepoint.com:/sites/Adsero:/lists')
      .filter('displayName eq "Documents"')
      .select('name,webUrl')
      .get();
    res.json(libraries.value);
  } catch (error) {
    res.status(500).json({ error: 'Error fetching libraries' });
  }
});

// Start Server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
