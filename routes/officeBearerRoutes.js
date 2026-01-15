const express = require('express');
const router = express.Router();
const upload = require('../middleware/upload');
const { uploadExcel, getAllOfficeBearers } = require('../controllers/officeBearerController');

// Upload Excel file and create office bearers
router.post('/upload', upload.single('excelFile'), uploadExcel);

// Get all office bearers
router.get('/', getAllOfficeBearers);

module.exports = router;
