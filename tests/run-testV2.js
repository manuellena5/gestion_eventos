// Corre testV2() y borrarDatosDePrueba() de Code.gs con un simulador de Apps Script.
// Uso: node tests/run-testV2.js
const { makeEnv } = require('./gas-mock');
const env = makeEnv();
env.run('testV2()');
console.log(env.logs.join('\n'));
env.logs.length = 0;
env.run('borrarDatosDePrueba()');
console.log(env.logs.join('\n'));
const fallo = /FALLA/.test(JSON.stringify(env.logs)) ;
process.exit(fallo ? 1 : 0);
