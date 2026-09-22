const { createClient } = require('@supabase/supabase-js');

// SUPABASE_URL is not a secret (it's already public in the old code.js), but
// reading it from the environment too lets it be overridden per-deployment
// without a code change.
const SUPABASE_URL = process.env.SUPABASE_URL || 'https://fcloggepmjmkbwkemtsy.supabase.co';
const SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;

let client = null;

// Server-only client, authorized as service_role — bypasses RLS entirely.
// Every handler in lib/handlers.js is responsible for its own authorization
// checks (same trust model as the old Apps Script backend), since RLS can't
// do that job here.
function getAdminClient() {
  if (!SERVICE_ROLE_KEY) {
    throw new Error('SUPABASE_SERVICE_ROLE_KEY environment variable is not set.');
  }
  if (!client) {
    client = createClient(SUPABASE_URL, SERVICE_ROLE_KEY, {
      auth: { autoRefreshToken: false, persistSession: false }
    });
  }
  return client;
}

module.exports = { getAdminClient, SUPABASE_URL };
