const express = require('express');
const cors = require('cors');
const xlsx = require('xlsx');
const bodyParser = require('body-parser');
const path = require('path');
const fs = require('fs-extra');
const morgan = require('morgan');

// Create the Express application
const app = express();
const PORT = 3001;

// Set up logging
const logStream = fs.createWriteStream(path.join(__dirname, 'logs', 'server.log'), { flags: 'a' });
app.use(morgan('combined', { stream: logStream }));
app.use(morgan('dev')); // Console logging

// Middleware
app.use(cors());
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));

// Excel file paths
const SETTINGS_FILE = path.join(__dirname, 'excel', 'settings.xlsx');
const VEHICLES_FILE = path.join(__dirname, 'excel', 'vehicles.xlsx');
const CARTONS_FILE = path.join(__dirname, 'excel', 'cartons.xlsx');

// Ensure Excel files exist
function ensureExcelFilesExist() {
  // Create excel directory if it doesn't exist
  const excelDir = path.dirname(SETTINGS_FILE);
  if (!fs.existsSync(excelDir)) {
    fs.mkdirSync(excelDir, { recursive: true });
    console.log(`Created directory: ${excelDir}`);
  }

  // Check settings file
  if (!fs.existsSync(SETTINGS_FILE)) {
    console.log('Creating settings.xlsx file...');
    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.aoa_to_sheet([
      ['Setting', 'Value'],
      ['LoginPIN', '1234'] // Default PIN
    ]);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Settings');
    xlsx.writeFile(workbook, SETTINGS_FILE);
    console.log('settings.xlsx created successfully');
  }

  // Check vehicles file
  if (!fs.existsSync(VEHICLES_FILE)) {
    console.log('Creating vehicles.xlsx file...');
    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.aoa_to_sheet([
      ['ID', 'Name'],
      ['TRUCK-001', 'Delivery Van 1'],
      ['TRUCK-002', 'Delivery Van 2']
    ]);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Vehicles');
    xlsx.writeFile(workbook, VEHICLES_FILE);
    console.log('vehicles.xlsx created successfully');
  }

  // Check cartons file
  if (!fs.existsSync(CARTONS_FILE)) {
    console.log('Creating cartons.xlsx file...');
    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.aoa_to_sheet([
      ['ID', 'Status', 'VehicleID', 'DateScanned', 'DatePickedUp', 'DateDelivered', 'AdditionalData']
    ]);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Cartons');
    xlsx.writeFile(workbook, CARTONS_FILE);
    console.log('cartons.xlsx created successfully');
  }
}

// Helper function to read Excel file
function readExcelFile(filePath) {
  try {
    // Check if file exists
    if (!fs.existsSync(filePath)) {
      console.error(`File not found: ${filePath}`);
      return [];
    }

    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Convert to JSON
    const data = xlsx.utils.sheet_to_json(worksheet);
    
    // Debug output
    console.log(`Read ${data.length} rows from ${filePath}`);
    if (data.length > 0) {
      console.log('First row sample:', JSON.stringify(data[0]).substring(0, 100) + '...');
    }
    
    return data;
  } catch (error) {
    console.error(`Error reading ${filePath}:`, error);
    return [];
  }
}

// Helper function to write Excel file
function writeExcelFile(filePath, data) {
  try {
    // Create a backup before writing
    const backupPath = `${filePath}.backup`;
    if (fs.existsSync(filePath)) {
      fs.copyFileSync(filePath, backupPath);
    }

    const worksheet = xlsx.utils.json_to_sheet(data);
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    xlsx.writeFile(workbook, filePath);
    
    console.log(`Successfully wrote to ${filePath}`);
    return true;
  } catch (error) {
    console.error(`Error writing ${filePath}:`, error);
    
    // Try to restore from backup if write failed
    const backupPath = `${filePath}.backup`;
    if (fs.existsSync(backupPath)) {
      try {
        fs.copyFileSync(backupPath, filePath);
        console.log(`Restored ${filePath} from backup`);
      } catch (restoreError) {
        console.error(`Failed to restore from backup:`, restoreError);
      }
    }
    
    return false;
  }
}

// API Routes

// Verify login PIN
app.post('/api/login', (req, res) => {
  console.log('POST /api/login');
  const { pin } = req.body;
  
  if (!pin) {
    return res.status(400).json({ error: 'PIN is required' });
  }
  
  console.log('Reading settings file...');
  const settings = readExcelFile(SETTINGS_FILE);
  console.log('Settings data:', settings);
  
  // Find the login PIN setting
  const loginSetting = settings.find(s => s.Setting === 'LoginPIN');
  console.log('Login setting found:', loginSetting);
  
  if (!loginSetting) {
    // Fallback to hardcoded PIN if setting not found
    console.log('Login PIN setting not found, using fallback PIN');
    if (pin === '1234') {
      return res.json({ success: true });
    } else {
      return res.status(401).json({ error: 'Invalid PIN' });
    }
  }
  
  console.log('Comparing PIN:', pin, 'with stored PIN:', loginSetting.Value.toString());
  if (pin === loginSetting.Value.toString()) {
    res.json({ success: true });
  } else {
    res.status(401).json({ error: 'Invalid PIN' });
  }
});

// Get all vehicles
app.get('/api/vehicles', (req, res) => {
  console.log('GET /api/vehicles');
  const vehicles = readExcelFile(VEHICLES_FILE);
  res.json(vehicles);
});

// Add a new vehicle
app.post('/api/vehicles', (req, res) => {
  console.log('POST /api/vehicles', req.body);
  const { id, name } = req.body;
  
  if (!id || !name) {
    return res.status(400).json({ error: 'Vehicle ID and Name are required' });
  }
  
  const vehicles = readExcelFile(VEHICLES_FILE);
  
  // Check if vehicle already exists
  if (vehicles.some(vehicle => vehicle.ID === id)) {
    return res.status(400).json({ error: 'Vehicle already exists' });
  }
  
  // Add new vehicle
  const newVehicle = {
    ID: id,
    Name: name
  };
  
  vehicles.push(newVehicle);
  
  if (writeExcelFile(VEHICLES_FILE, vehicles)) {
    res.status(201).json(newVehicle);
  } else {
    res.status(500).json({ error: 'Failed to save vehicle' });
  }
});

// Get all cartons
app.get('/api/cartons', (req, res) => {
  console.log('GET /api/cartons');
  const cartons = readExcelFile(CARTONS_FILE);
  res.json(cartons);
});

// Add multiple cartons
app.post('/api/cartons/batch', (req, res) => {
  console.log('POST /api/cartons/batch', req.body);
  const { cartons: newCartons, cartonsData } = req.body;
  
  if (!newCartons || !Array.isArray(newCartons) || newCartons.length === 0) {
    return res.status(400).json({ error: 'Cartons array is required' });
  }
  
  const cartons = readExcelFile(CARTONS_FILE);
  const timestamp = new Date().toISOString();
  
  // Process each carton
  const addedCartons = [];
  const existingCartons = [];
  
  newCartons.forEach(id => {
    // Check if carton already exists
    if (cartons.some(carton => carton.ID === id)) {
      existingCartons.push(id);
      return;
    }
    
    // Find additional data if available
    let additionalData = {};
    if (cartonsData) {
      const cartonData = cartonsData.find(c => c.id === id);
      if (cartonData && cartonData.additionalData) {
        additionalData = cartonData.additionalData;
      }
    }
    
    // Convert additional data to string for Excel
    let additionalDataStr = '';
    if (Object.keys(additionalData).length > 0) {
      additionalDataStr = JSON.stringify(additionalData);
    }
    
    // Add new carton
    const newCarton = {
      ID: id,
      Status: 'unassigned',
      VehicleID: '',
      DateScanned: timestamp,
      DatePickedUp: '',
      DateDelivered: '',
      AdditionalData: additionalDataStr
    };
    
    cartons.push(newCarton);
    addedCartons.push(id);
  });
  
  if (addedCartons.length > 0) {
    if (writeExcelFile(CARTONS_FILE, cartons)) {
      res.status(201).json({ 
        success: true, 
        added: addedCartons, 
        existing: existingCartons,
        count: addedCartons.length
      });
    } else {
      res.status(500).json({ error: 'Failed to save cartons' });
    }
  } else {
    res.status(400).json({ 
      success: false, 
      error: 'All cartons already exist', 
      existing: existingCartons 
    });
  }
});

// Add a new carton
app.post('/api/cartons', (req, res) => {
  console.log('POST /api/cartons', req.body);
  const { id, status = 'unassigned', vehicleId = '', additionalData = {} } = req.body;
  
  if (!id) {
    return res.status(400).json({ error: 'Carton ID is required' });
  }
  
  const cartons = readExcelFile(CARTONS_FILE);
  
  // Check if carton already exists
  if (cartons.some(carton => carton.ID === id)) {
    return res.status(400).json({ error: 'Carton already exists' });
  }
  
  // Convert additional data to string for Excel
  let additionalDataStr = '';
  if (Object.keys(additionalData).length > 0) {
    additionalDataStr = JSON.stringify(additionalData);
  }
  
  // Add new carton
  const newCarton = {
    ID: id,
    Status: status,
    VehicleID: vehicleId,
    DateScanned: new Date().toISOString(),
    DatePickedUp: '',
    DateDelivered: '',
    AdditionalData: additionalDataStr
  };
  
  cartons.push(newCarton);
  
  if (writeExcelFile(CARTONS_FILE, cartons)) {
    res.status(201).json(newCarton);
  } else {
    res.status(500).json({ error: 'Failed to save carton' });
  }
});

// Assign cartons to vehicle
app.post('/api/assign', (req, res) => {
  console.log('POST /api/assign', req.body);
  const { cartonIds, vehicleId } = req.body;
  
  if (!cartonIds || !vehicleId) {
    return res.status(400).json({ error: 'Carton IDs and Vehicle ID are required' });
  }
  
  const cartons = readExcelFile(CARTONS_FILE);
  const vehicles = readExcelFile(VEHICLES_FILE);
  
  // Check if vehicle exists
  if (!vehicles.some(vehicle => vehicle.ID === vehicleId)) {
    return res.status(400).json({ error: 'Vehicle not found' });
  }
  
  // Update cartons
  const updatedCartons = cartons.map(carton => {
    if (cartonIds.includes(carton.ID)) {
      return { ...carton, Status: 'assigned', VehicleID: vehicleId };
    }
    return carton;
  });
  
  if (writeExcelFile(CARTONS_FILE, updatedCartons)) {
    res.json({ success: true, count: cartonIds.length });
  } else {
    res.status(500).json({ error: 'Failed to assign cartons' });
  }
});

// Confirm carton pickup by driver
app.post('/api/pickup', (req, res) => {
  console.log('POST /api/pickup', req.body);
  const { cartonId, vehicleId } = req.body;
  
  if (!cartonId) {
    return res.status(400).json({ error: 'Carton ID is required' });
  }
  
  const cartons = readExcelFile(CARTONS_FILE);
  
  // Find the carton
  const carton = cartons.find(c => c.ID === cartonId);
  
  if (!carton) {
    return res.status(404).json({ error: 'Carton not found' });
  }
  
  // Check if carton is assigned to the correct vehicle
  if (vehicleId && carton.VehicleID !== vehicleId) {
    return res.status(400).json({ 
      error: 'Carton is not assigned to this vehicle',
      assignedTo: carton.VehicleID
    });
  }
  
  // Check if carton is in the correct state
  if (carton.Status !== 'assigned') {
    return res.status(400).json({ 
      error: `Carton cannot be picked up because it is ${carton.Status}`,
      status: carton.Status
    });
  }
  
  // Update carton status
  const updatedCartons = cartons.map(c => {
    if (c.ID === cartonId) {
      return { 
        ...c, 
        Status: 'picked_up',
        DatePickedUp: new Date().toISOString()
      };
    }
    return c;
  });
  
  if (writeExcelFile(CARTONS_FILE, updatedCartons)) {
    res.json({ 
      success: true, 
      cartonId,
      status: 'picked_up'
    });
  } else {
    res.status(500).json({ error: 'Failed to update carton status' });
  }
});

// Confirm delivery
app.post('/api/deliver', (req, res) => {
  console.log('POST /api/deliver', req.body);
  const { cartonId } = req.body;
  
  if (!cartonId) {
    return res.status(400).json({ error: 'Carton ID is required' });
  }
  
  const cartons = readExcelFile(CARTONS_FILE);
  
  // Find the carton
  const carton = cartons.find(c => c.ID === cartonId);
  
  if (!carton) {
    return res.status(404).json({ error: 'Carton not found' });
  }
  
  // Check if carton is in the correct state
  if (carton.Status !== 'picked_up') {
    return res.status(400).json({ 
      error: `Carton must be picked up before delivery`,
      status: carton.Status
    });
  }
  
  // Update carton status
  const updatedCartons = cartons.map(c => {
    if (c.ID === cartonId) {
      return { 
        ...c, 
        Status: 'delivered',
        DateDelivered: new Date().toISOString()
      };
    }
    return c;
  });
  
  if (writeExcelFile(CARTONS_FILE, updatedCartons)) {
    res.json({ 
      success: true, 
      cartonId,
      status: 'delivered'
    });
  } else {
    res.status(500).json({ error: 'Failed to update carton status' });
  }
});

// Health check endpoint
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', time: new Date().toISOString() });
});

// Serve the React app
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Start server
app.listen(PORT, '0.0.0.0', () => {
  console.log(`=================================================`);
  console.log(`Delivery Management System Server`);
  console.log(`=================================================`);
  console.log(`Server running at:`);
  console.log(`- Local:   http://localhost:${PORT}`);
  
  // Try to get the local IP address
  try {
    const { networkInterfaces } = require('os');
    const nets = networkInterfaces();
    for (const name of Object.keys(nets)) {
      for (const net of nets[name]) {
        // Skip internal and non-IPv4 addresses
        if (net.family === 'IPv4' && !net.internal) {
          console.log(`- Network: http://${net.address}:${PORT}`);
        }
      }
    }
  } catch (err) {
    console.log('Could not determine network IP address');
  }
  
  console.log(`=================================================`);
  console.log(`Excel files location:`);
  console.log(`- Settings: ${SETTINGS_FILE}`);
  console.log(`- Vehicles: ${VEHICLES_FILE}`);
  console.log(`- Cartons: ${CARTONS_FILE}`);
  console.log(`=================================================`);
  
  // Ensure Excel files exist
  ensureExcelFilesExist();
});