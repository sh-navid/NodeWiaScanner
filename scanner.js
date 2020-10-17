'use strict';

let run = () => {
    const spawn = require('child_process').spawnSync;
    const vbs = spawn('cscript.exe', ['scan.vbs', 'one']);

    console.log(`stderr: ${vbs.stderr.toString()}`);
    console.log(`stdout: ${vbs.stdout.toString()}`);
    console.log(`status: ${vbs.status}`);
}

run();