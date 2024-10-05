// ipcMainHandlers.js
import { ipcMain } from 'electron';
import { spawn } from 'child_process';
import { stderr, stdout } from 'process';

export function initializeIpcMainHandlers() {
    ipcMain.on('excel-opener-1', (event) => {
        const python = spawn('python', [
            'src/Scripts/Python/main/excel-opener-1.py',
        ]);

        python.stdout.on('data', (data) => {
            console.log(`Output: ${data}`);
        });
    });

    ipcMain.on('excel-opener-2', (event) => {
        const python = spawn('python', [
            'src/Scripts/Python/main/excel-opener-2.py',
        ]);

        python.stdout.on('data', (data) => {
            console.log(`Output: ${data}`);
        });
    });

    ipcMain.on('resize-excel', (event) => {
        const python = spawn('python', [
            'src/Scripts/Python/main/resize-excel.py',
        ]);

        python.stdout.on('data', (data) => {
            console.log(`Output: ${data}`);
        });
    });

    ipcMain.on('active-window', (event) => {
        const python = spawn('python', [
            'src/Scripts/Python/main/active-window.py',
        ]);

        python.stdout.on('data', (data) => {
            console.log(`Output: ${data}`);
        });
    });

    ipcMain.on('excel-always-enable', (event) => {
        const python = spawn('python', [
            'src/Scripts/Python/main/enable-always-on-top.py',
        ]);
        python.stdout.on('data', (data) => {
            console.log(`Output: ${data}`);
        });
    });

    ipcMain.on('excel-always-disable', (event) => {
        const python = spawn('python', [
            'src/Scripts/Python/main/disable-always-on-top.py',
        ]);

        python.stdout.on('data', (data) => {
            console.log(`Output: ${data}`);
        });
    });

    ipcMain.on('close-excel', (event) => {
        const python = spawn('python', [
            'src/Scripts/Python/main/close-excel.py',
        ]);

        python.stdout.on('data', (data) => {
            console.log(`Output: ${data}`);
        });
    });

    ipcMain.on('project-1', (event) => {
        const python = spawn('python', ['src/Scripts/Python/vba/Project 1.py']);

        python.stdout.on('data', (data) => {
            console.log(`Output: ${data}`);
        });
    });

    ipcMain.on('project-2', (event) => {
        const python = spawn('python', ['src/Scripts/Python/vba/Project 2.py']);

        python.stdout.on('data', (data) => {
            console.log(`Output: ${data}`);
        });
    });

    ipcMain.on('project-valuation-1', (event) => {
        const python = spawn('python', [
            'src/Scripts/Python/Valuation/Project1.py',
        ]);

        python.stdout.on('data', (data) => {
            // Convert buffer to string and trim any extra new lines
            const output = data.toString().trim();
            console.log(`Output: ${output}`);

            // Send the output back to the renderer process
            event.reply('project-valuation-output-1', output);
        });
    });

    ipcMain.on('project-valuation-2', (event) => {
        const python = spawn('python', [
            'src/Scripts/Python/Valuation/Project2.py',
        ]);

        python.stdout.on('data', (data) => {
            // Convert buffer to string and trim any extra new lines
            const output = data.toString().trim();
            console.log(`Output: ${output}`);

            // Send the output back to the renderer process
            event.reply('project-valuation-output-2', output);
        });
    });
}
