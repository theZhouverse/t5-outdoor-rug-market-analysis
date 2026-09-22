import path from 'node:path';
import { spawn } from 'node:child_process';
import { pathToFileURL } from 'node:url';

const ROOT = process.cwd();
const PY = process.env.PARENT_PYTHON || path.join(process.env.USERPROFILE || '', '.cache', 'codex-runtimes', 'codex-primary-runtime', 'dependencies', 'python', 'python.exe');

function runPython(file, args = []) {
  return new Promise((resolve, reject) => {
    const child = spawn(PY, ['-u', file, ...args], { cwd: ROOT, stdio: 'inherit', windowsHide: true, env: { ...process.env, PYTHONIOENCODING: 'utf-8' } });
    child.on('error', reject);
    child.on('exit', (code) => code === 0 ? resolve() : reject(new Error(`${file} failed with exit code ${code}`)));
  });
}

export async function buildParentDelivery() {
  const model = spawn(process.execPath, ['src/build_parent_source_model.mjs'], { cwd: ROOT, stdio: 'inherit', windowsHide: true });
  await new Promise((resolve, reject) => { model.on('error', reject); model.on('exit', (code) => code === 0 ? resolve() : reject(new Error(`model build failed with exit code ${code}`))); });
  await runPython('src/build_parent_source_xlsx.py');
  const child = spawn(process.execPath, ['src/build_parent_source_html.mjs'], { cwd: ROOT, stdio: 'inherit', windowsHide: true });
  await new Promise((resolve, reject) => { child.on('error', reject); child.on('exit', (code) => code === 0 ? resolve() : reject(new Error(`HTML build failed with exit code ${code}`))); });
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) await buildParentDelivery();
