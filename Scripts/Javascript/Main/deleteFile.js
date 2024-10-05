const fs = require('fs');
const path = require('path');

// Specify the file path
const filePath = path.join(__dirname, 'new-folder');

// Delete the file
fs.unlink(filePath, (err) => {
    if (err) {
        return console.error(`Error deleting file: ${err}`);
    }
    console.log('File deleted successfully!');
});
