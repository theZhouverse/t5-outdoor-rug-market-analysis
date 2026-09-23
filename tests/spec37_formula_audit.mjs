import path from 'node:path';
import { spawn } from 'node:child_process';

const root = process.cwd();
const python = process.env.PARENT_PYTHON || path.join(process.env.USERPROFILE || '', '.cache', 'codex-runtimes', 'codex-primary-runtime', 'dependencies', 'python', 'python.exe');
const child = spawn(python, ['-u', 'tests/spec37_formula_audit.py'], {
  cwd: root,
  stdio: 'inherit',
  windowsHide: true,
  env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
});
child.on('error', (error) => { console.error(error); process.exitCode = 1; });
child.on('exit', (code) => { process.exitCode = code ?? 1; });
