const path = require('path');
const Service = require('node-windows').Service;

// Use an absolute path for the script
const serverPath = path.join(__dirname, 'server.js');

// Create a new service object
const svc = new Service({
  name: 'ExpressGEOTAG',
  description: 'ExpressGEOTAG',
  script: serverPath
});

// Listen for the "install" event, which indicates the service is installed
svc.on('install', function() {
  console.log('Service installed.');
  svc.start();
});

// In case the service is already installed
svc.on('alreadyinstalled', function() {
  console.log('Service is already installed.');
});

// If there are any errors
svc.on('error', function(err) {
  console.error('Service error:', err);
});

// If the installation process has been initiated elsewhere, uninstall it first
svc.on('uninstall', function() {
  console.log('Uninstall complete.');
  console.log('The service exists:', svc.exists);
});

// Install the service.
svc.uninstall();
