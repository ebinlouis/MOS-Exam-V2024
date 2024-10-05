const fs = require('fs');
const path = require('path');

// Specify the directory name
const directoryPath = path.join(__dirname, 'new-folder');

// Create the directory
fs.mkdir(directoryPath, { recursive: true }, (err) => {
    if (err) {
        return console.error(`Error creating directory: ${err}`);
    }
    console.log('Directory created successfully!');
});
