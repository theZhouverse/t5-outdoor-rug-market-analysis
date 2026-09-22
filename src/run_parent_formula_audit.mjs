import path from 'node:path';
import { spawn } from 'node:child_process';

const root=process.cwd();
const python=process.env.PARENT_PYTHON || path.join(process.env.USERPROFILE || '', '.cache', 'codex-runtimes', 'codex-primary-runtime', 'dependencies', 'python', 'python.exe');
const child=spawn(python,['-u','tests/parent_formula_audit.py'],{cwd:root,stdio:'inherit',windowsHide:true,env:{...process.env,PYTHONIOENCODING:'utf-8'}});
await new Promise((resolve,reject)=>{
  child.on('error',reject);
  child.on('exit',code=>code===0?resolve():reject(new Error(`formula audit failed with exit code ${code}`)));
});
