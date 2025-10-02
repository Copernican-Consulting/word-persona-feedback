const fs = require('fs'); const path = require('path');
const baseUrl = process.env.BASE_URL || (process.env.npm_package_config_baseUrl) || '';
if (!baseUrl) { console.error('ERROR: Please set BASE_URL env var or package.json config.baseUrl'); process.exit(1); }
const tpl = fs.readFileSync(path.join(__dirname, '..', 'manifest.template.xml'), 'utf8');
const xml = tpl.replace(/__BASE_URL__/g, baseUrl.replace(/\/$/, ''));
const out = path.join(__dirname, '..', 'manifest.dev.xml');
fs.writeFileSync(out, xml, 'utf8');
console.log('Wrote', out, 'with BASE_URL =', baseUrl);