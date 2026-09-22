/**
 * Single RPC endpoint the frontend's callServer(fnName, payload) posts to
 * — see index.html. Mirrors the old google.script.run.withSuccessHandler()
 * [...] [fnName](payload) pattern closely enough that only callServer()
 * itself needed to change, not any of its ~20 call sites.
 */

const handlers = require('../lib/handlers');

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    res.status(405).json({ success: false, message: 'Method not allowed.' });
    return;
  }

  const { fn, payload } = req.body || {};
  const handler = Object.prototype.hasOwnProperty.call(handlers, fn) ? handlers[fn] : null;

  if (!handler) {
    res.status(400).json({ success: false, message: 'Unknown function: ' + fn });
    return;
  }

  try {
    const result = await handler(payload);
    res.status(200).json(result);
  } catch (err) {
    res.status(500).json({ success: false, message: 'Server error: ' + err.message });
  }
};
