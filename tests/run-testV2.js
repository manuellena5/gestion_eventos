// Corre testV2() y borrarDatosDePrueba() de Code.gs con un simulador de Apps Script.
// Uso: node tests/run-testV2.js
const { makeEnv } = require('./gas-mock');
let fallo = false;

// 1) Corrida normal
let env = makeEnv();
env.run('testV2()');
console.log(env.logs.join('\n'));
fallo = fallo || env.logs.some(l => l.indexOf('FALLA') === 0);
const ops = env.ss.getSheets().reduce((s, sh) => s + sh.ops, 0);
console.log(`Accesos a hojas: ${ops} · pedidos de zona horaria: ${env.stats.tz}`);
env.logs.length = 0;
env.run('borrarDatosDePrueba()');
console.log(env.logs.slice(-1).join('\n'));

// 2) Con errores temporales de Google: tiene que reintentar y pasar igual
env = makeEnv();
env.stats.failEvery = 7; // 1 de cada 7 lecturas falla
env.run('testV2()');
const f2 = env.logs.filter(l => l.indexOf('FALLA') === 0);
console.log(`Con ${env.stats.failed} errores temporales inyectados: ${f2.length ? f2.length + ' fallas\n' + f2.join('\n') : 'todas las pruebas pasaron'}`);
fallo = fallo || f2.length > 0;
process.exit(fallo ? 1 : 0);
