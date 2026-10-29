// Builds the static dashboard site in public/ from dashboard/page.html (the same page used on claude.ai).
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const shell = fs.readFileSync(path.join(root, 'dashboard', 'shell.html'), 'utf8');
const page = fs.readFileSync(path.join(root, 'dashboard', 'page.html'), 'utf8').replace(/<meta name="dash-one-writer">\r?\n/, '');
fs.mkdirSync(path.join(root, 'public'), { recursive: true });
fs.writeFileSync(path.join(root, 'public', 'index.html'), shell.replace('<!-- PAGE -->', () => page));
fs.copyFileSync(path.join(root, 'dashboard', 'dash-shim.js'), path.join(root, 'public', 'dash-shim.js'));
fs.copyFileSync(path.join(root, 'dashboard', 'reports.js'), path.join(root, 'public', 'reports.js'));
fs.copyFileSync(path.join(root, 'dashboard', 'departments.js'), path.join(root, 'public', 'departments.js'));
console.log('built public/index.html');
