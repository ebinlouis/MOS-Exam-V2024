const { spawn } = require('child_process');

export const excelOpenerOne = () => {
    const pythonProcess = spawn('python', [
        'src/Scripts/Python/main/excel-opener-1.py',
    ]);

    pythonProcess.stdout.on('data', (data) => {
        // setOutput(data.toString());
        console.log(`Python output: ${data.toString()}`);
    });
};

export const excelResize = () => {
    const pythonProcess = spawn('python', [
        'src/Scripts/Python/main/resize-excel.py',
    ]);

    pythonProcess.stdout.on('data', (data) => {
        // setOutput(data.toString());
        console.log(`Python output: ${data.toString()}`);
    });
};

export const excelAlways = () => {
    const pythonProcess = spawn('python', [
        'src/Scripts/Python/main/enable-always-on-top.py',
    ]);

    pythonProcess.stdout.on('data', (data) => {
        // setOutput(data.toString());
        console.log(`Python output: ${data.toString()}`);
    });
};

export const excelOpenerTwo = () => {
    const pythonProcess = spawn('python', [
        'src/Scripts/Python/main/excel-opener-2.py',
    ]);

    pythonProcess.stdout.on('data', (data) => {
        // setOutput(data.toString());
        console.log(`Python output: ${data.toString()}`);
    });
};

export const projectValuationOne = (setTaskOne) => {
    const pythonProcess = spawn('python', [
        'src/Scripts/Python/Valuation/Project1.py',
    ]);

    pythonProcess.stdout.on('data', (data) => {
        setTaskOne(data.toString());
        console.log(`Python output: ${data.toString()}`);
    });
};

export const projectValuationTwo = (setTaskTwo) => {
    const pythonProcess = spawn('python', [
        'src/Scripts/Python/Valuation/Project2.py',
    ]);

    pythonProcess.stdout.on('data', (data) => {
        setTaskTwo(data.toString());
        console.log(`Python output: ${data.toString()}`);
    });
};

export const projectOne = () => {
    const pythonProcess = spawn('python', [
        'src/Scripts/Python/vba/Project 1.py',
    ]);

    pythonProcess.stdout.on('data', (data) => {
        // setTaskTwo(data.toString());
        console.log(`Python output: ${data.toString()}`);
    });
};

export const projectTwo = () => {
    const pythonProcess = spawn('python', [
        'src/Scripts/Python/vba/Project 2.py',
    ]);

    pythonProcess.stdout.on('data', (data) => {
        // setTaskTwo(data.toString());
        console.log(`Python output: ${data.toString()}`);
    });
};

export const closeExcel = () => {
    const pythonProcess = spawn('python', [
        'src/Scripts/Python/main/close-excel.py',
    ]);

    pythonProcess.stdout.on('data', (data) => {
        // setTaskTwo(data.toString());
        console.log(`Python output: ${data.toString()}`);
    });
};
