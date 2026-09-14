import {spawnSync} from 'node:child_process';
import path from 'node:path';
const py=process.env.SPEC21_PYTHON || path.join(process.env.USERPROFILE || '', '.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe');
const jobs=process.argv.includes('--inputs')?[[py,['-u','tests/spec21_formula_inputs.py']]]:process.argv.includes('--delivery')?[[py,['-u','tests/spec21_delivery_audit.py']]]:[[process.execPath,['tests/spec21_model_audit.mjs']],[process.execPath,['tests/spec2_audit.mjs']],[py,['-u','tests/spec21_source_audit.py']],[py,['-u','tests/spec21_delivery_audit.py']]];
for(const [exe,args] of jobs){const result=spawnSync(exe,args,{stdio:'inherit',windowsHide:true,env:{...process.env,PYTHONIOENCODING:'utf-8'}});if(result.error)throw result.error;if(result.status!==0)process.exit(result.status || 1);}
