// Polyfills window.storage using the browser's localStorage,
// matching the get/set contract that moneytracker.tsx already calls.
// This lets the app run standalone (Vercel/Netlify/Replit/etc.)
// without any changes to moneytracker.tsx itself.

function keyFor(key, shared) {
  return `moneytracker:${shared ? 'shared' : 'personal'}:${key}`;
}

window.storage = {
  async get(key, shared = false) {
    const raw = localStorage.getItem(keyFor(key, shared));
    if (raw === null) return null;
    return { key, value: raw, shared };
  },
  async set(key, value, shared = false) {
    localStorage.setItem(keyFor(key, shared), value);
    return { key, value, shared };
  },
  async delete(key, shared = false) {
    const existed = localStorage.getItem(keyFor(key, shared)) !== null;
    localStorage.removeItem(keyFor(key, shared));
    return { key, deleted: existed, shared };
  },
  async list(prefix = '', shared = false) {
    const scope = `moneytracker:${shared ? 'shared' : 'personal'}:`;
    const keys = [];
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (k && k.startsWith(scope + prefix)) keys.push(k.slice(scope.length));
    }
    return { keys, prefix, shared };
  },
};
