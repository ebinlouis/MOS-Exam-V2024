const fs = require('fs');
const path = require('path');

// Specify the directory path
const directoryPath = path.join(__dirname, 'new-folder');

// Delete the directory
fs.rmdir(directoryPath, (err) => {
    if (err) {
        return console.error(`Error deleting directory: ${err}`);
    }
    console.log('Directory deleted successfully!');
});
