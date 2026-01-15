const OfficeBearer = require('../models/OfficeBearer');
const { validateExcelData } = require('../utils/excelValidator');

/**
 * Upload Excel file and create office bearers
 */
async function uploadExcel(req, res) {
  try {
    // Check if file is uploaded
    if (!req.file) {
      return res.status(400).json({
        success: false,
        error: 'No file uploaded',
        message: 'Please upload an Excel file'
      });
    }

    // Validate Excel data
    const validation = validateExcelData(req.file.buffer);

    if (!validation.isValid) {
      return res.status(400).json({
        success: false,
        error: 'Validation failed',
        message: 'Excel file validation failed',
        errors: validation.errors
      });
    }

    const officeBearerData = validation.data;

    // Check for duplicate emails in the Excel file itself
    const emailSet = new Set();
    const duplicateEmails = [];
    
    for (let i = 0; i < officeBearerData.length; i++) {
      const email = officeBearerData[i].email.toLowerCase();
      if (emailSet.has(email)) {
        duplicateEmails.push({
          row: i + 2,
          email: officeBearerData[i].email,
          message: `Row ${i + 2}: Duplicate email "${officeBearerData[i].email}" found in Excel file`
        });
      } else {
        emailSet.add(email);
      }
    }

    if (duplicateEmails.length > 0) {
      return res.status(400).json({
        success: false,
        error: 'Duplicate emails in file',
        message: 'Duplicate email addresses found in the Excel file',
        errors: duplicateEmails
      });
    }

    // Check for existing emails in database
    const emails = officeBearerData.map(item => item.email.toLowerCase());
    const existingBearers = await OfficeBearer.findAll({
      where: {
        email: emails
      },
      attributes: ['email']
    });

    const existingEmails = existingBearers.map(bearer => bearer.email.toLowerCase());
    const conflictingEmails = officeBearerData.filter(item => 
      existingEmails.includes(item.email.toLowerCase())
    );

    if (conflictingEmails.length > 0) {
      return res.status(400).json({
        success: false,
        error: 'Duplicate emails in database',
        message: 'Some email addresses already exist in the database',
        errors: conflictingEmails.map(item => ({
          email: item.email,
          message: `Email "${item.email}" already exists in the database`
        }))
      });
    }

    // Create office bearers in database
    const createdBearers = [];
    const creationErrors = [];

    // Use transaction for atomicity
    const transaction = await OfficeBearer.sequelize.transaction();

    try {
      for (const data of officeBearerData) {
        try {
          const bearer = await OfficeBearer.create(data, { transaction });
          createdBearers.push(bearer);
        } catch (error) {
          creationErrors.push({
            email: data.email,
            message: `Failed to create office bearer: ${error.message}`
          });
        }
      }

      if (creationErrors.length > 0) {
        await transaction.rollback();
        return res.status(500).json({
          success: false,
          error: 'Creation failed',
          message: 'Some office bearers could not be created',
          errors: creationErrors
        });
      }

      await transaction.commit();

      return res.status(201).json({
        success: true,
        message: `Successfully created ${createdBearers.length} office bearer(s)`,
        count: createdBearers.length,
        data: createdBearers.map(bearer => ({
          id: bearer.id,
          name: bearer.name,
          email: bearer.email,
          phone: bearer.phone,
          position: bearer.position,
          department: bearer.department,
          address: bearer.address
        }))
      });
    } catch (error) {
      await transaction.rollback();
      throw error;
    }
  } catch (error) {
    console.error('Error uploading Excel file:', error);
    return res.status(500).json({
      success: false,
      error: 'Server error',
      message: 'An error occurred while processing the Excel file',
      details: error.message
    });
  }
}

/**
 * Get all office bearers
 */
async function getAllOfficeBearers(req, res) {
  try {
    const bearers = await OfficeBearer.findAll({
      order: [['createdAt', 'DESC']]
    });
    
    return res.status(200).json({
      success: true,
      count: bearers.length,
      data: bearers
    });
  } catch (error) {
    console.error('Error fetching office bearers:', error);
    return res.status(500).json({
      success: false,
      error: 'Server error',
      message: 'Failed to fetch office bearers'
    });
  }
}

module.exports = {
  uploadExcel,
  getAllOfficeBearers
};
