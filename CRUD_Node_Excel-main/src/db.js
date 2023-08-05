const mongoose = require('mongoose');

const MONGO_URL = 'mongodb://0.0.0.0:27017/DATA';

mongoose.connect(MONGO_URL, {
  useNewUrlParser: true,
  useUnifiedTopology: true,
});

const db = mongoose.connection;

db.on('error', (error) => {
  console.error('MongoDB connection error:', error);
});

db.once('open', () => {
  console.log('Connected to MongoDB successfully!');
});

module.exports = db;
