import {test} from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import vm from 'node:vm';

test('real HTML IDs and module dependencies initialize one file and sheet handler', async () => {
  const html = await readFile(new URL('../ui/index.html', import.meta.url), 'utf8');
  let app = await readFile(new URL('../ui/app.js', import.meta.url), 'utf8');
  const imports = [...app.matchAll(/^import \{([^}]+)\} from '([^']+)';/gm)];
  const dependencies = {};
  for (const [, names, path] of imports) {
    const module = await import(new URL('../ui/' + path, import.meta.url));
    for (const name of names.split(',').map(n => n.trim())) {
      assert.equal(typeof module[name], 'function', name);
      dependencies[name] = module[name];
    }
  }
  app = app.replace(/^import .*;\r?\n/gm, '');
  function node() {
    return {value: '', hidden: false, style: {setProperty() {}}, dataset: {}, listeners: {},
      classList: {add() {}, remove() {}, toggle() {}, contains() {return false;}},
      addEventListener(type, handler) {(this.listeners[type] ??= new Set()).add(handler);},
      removeEventListener(type, handler) {this.listeners[type]?.delete(handler);},
      setAttribute() {}, removeAttribute() {}, getAttribute() {return null;},
      querySelector() {return null;}, querySelectorAll() {return [];}, appendChild() {},
      getBoundingClientRect() {return {width: 255};}
    };
  }
  const nodes = new Map([...html.matchAll(/id="([^"]+)"/g)].map(m => [m[1], node()]));
  const document = {getElementById: id => nodes.get(id) ?? null,
    querySelector: () => node(), querySelectorAll: () => [],
    addEventListener() {}, createElement: node, body: node()};
  const context = { ...dependencies, document, window: {addEventListener() {}},
    api: () => new Promise(() => {}),
    localStorage: {getItem() {return null;}, setItem() {}},
    Intl, Date, URLSearchParams, AbortController, setTimeout, clearTimeout,
    setInterval, clearInterval, location: {search: ''}, history: {replaceState() {}},
    fetch: () => new Promise(() => {})
  };
  vm.runInNewContext(app, context, {filename: 'ui/app.js'});
  assert.equal(nodes.get('family-list').listeners.click.size, 1);
  assert.equal(nodes.get('sheet-tabs').listeners.click.size, 1);
});
