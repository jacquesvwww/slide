/**
 * SLIDE — backend handlers, ported from the old Apps Script / Google
 * Sheets prototype (see code.js) onto Supabase.
 *
 * Every function here mirrors a function that used to live in code.js,
 * same name, same payload shape, same { success, message, ... } response
 * shape — so the frontend (index.html) barely had to change. The one
 * structural difference: profiles/shifts columns are the real Postgres
 * schema (snake_case), and shift reads go through the `shift_details`
 * view, which joins in owner/barista names from `profiles` rather than
 * storing them as duplicated text columns like the old Sheets version did.
 *
 * Every handler runs against the service_role client (see
 * supabaseAdmin.js), which bypasses Row Level Security entirely. That's
 * deliberate, not an oversight: the RLS policies on `shifts` only allow a
 * barista to update a row they're *already* attached to, which can't cover
 * first-time actions like claiming an open shift or making an initial rate
 * offer. So, same as the old Apps Script backend, every handler below is
 * responsible for its own authorization checks — the database doesn't do
 * it for us here.
 */

const { getAdminClient } = require('./supabaseAdmin');
const { geocodeAddress, getTravelTimes } = require('./locationiq');

// UK National Living Wage (21+), effective 1 April 2026. Update each April.
// https://www.gov.uk/national-minimum-wage-rates
const MINIMUM_WAGE = 12.71;

// Supabase Storage bucket used for logo/café photo uploads (replaces the
// old Google Drive folder — see uploadImage()).
const STORAGE_BUCKET = 'cafe-images';

// Vercel serverless functions cap request bodies at 4.5MB. A base64-encoded
// image is ~33% larger than the raw bytes, so this needs real headroom
// under that limit (unlike the old 5MB Apps Script cap, which had no such
// ceiling to worry about).
const MAX_IMAGE_SIZE_BYTES = 3 * 1024 * 1024; // 3 MB

/* ================================================================
   Shape helpers — map Postgres rows to the same camelCase objects
   the frontend already expects.
   ================================================================ */

function profileToSafeUser(row) {
  if (!row) return null;
  const ratingCount = Number(row.rating_count) || 0;
  const ratingTotal = Number(row.rating_total) || 0;
  return {
    userId: row.id,
    role: row.role,
    fullName: row.full_name,
    email: row.email,
    cafeName: row.cafe_name || '',
    phone: row.phone || '',
    availabilityType: row.availability_type || '',
    locationResolved: row.lat != null && row.lng != null,
    createdAt: row.created_at,
    bio: row.bio || '',
    address: row.address || '',
    logoUrl: row.logo_url || '',
    instagramUrl: row.instagram_url || '',
    facebookUrl: row.facebook_url || '',
    websiteUrl: row.website_url || '',
    cafeImages: row.cafe_images || [],
    ratingCount: ratingCount,
    ratingAverage: ratingCount > 0 ? Math.round((ratingTotal / ratingCount) * 10) / 10 : null
  };
}

// Postgres `time without time zone` comes back from PostgREST as
// "HH:MM:SS" — <input type="time"> requires zero-padded "HH:MM", so this
// just truncates rather than reformatting.
function formatTime_(val) {
  return val ? String(val).slice(0, 5) : '';
}

function shiftDetailsToObject(row) {
  return {
    shiftId: row.id,
    ownerUserId: row.owner_id,
    cafeName: row.cafe_name || row.owner_cafe_name || '',
    date: row.shift_date,
    startTime: formatTime_(row.start_time),
    endTime: formatTime_(row.end_time),
    rate: Number(row.rate) || 0,
    skills: row.skills || [],
    busyLevel: Number(row.busy_level) || 3,
    lunchIncluded: !!row.lunch_included,
    breakMinutes: Number(row.break_minutes) || 0,
    breakNotes: row.break_notes || '',
    status: row.status,
    assignedBaristaName: row.assigned_barista_name || '',
    assignedBaristaUserId: row.assigned_barista_id || '',
    createdAt: row.created_at,
    proposedRate: row.proposed_rate != null ? Number(row.proposed_rate) : null,
    proposedByName: row.proposed_by_name || '',
    proposedByUserId: row.proposed_by_id || '',
    negotiationStatus: row.negotiation_status || '',
    cancelledBy: row.cancelled_by || '',
    repostedShiftId: row.reposted_shift_id || '',
    notes: row.notes || '',
    rateNegotiable: row.rate_negotiable !== false,
    baristaRatingOfCafe: row.barista_rating_of_cafe != null ? Number(row.barista_rating_of_cafe) : null,
    ownerRatingOfBarista: row.owner_rating_of_barista != null ? Number(row.owner_rating_of_barista) : null
  };
}

async function fetchShiftDetails_(shiftId) {
  const db = getAdminClient();
  const { data, error } = await db.from('shift_details').select('*').eq('id', shiftId).maybeSingle();
  if (error) throw error;
  return data ? shiftDetailsToObject(data) : null;
}

/* ================================================================
   USERS — registration, login, profile
   ================================================================ */

async function registerUser(payload) {
  try {
    const db = getAdminClient();
    const email = String(payload.email || '').trim().toLowerCase();
    const password = String(payload.password || '');
    const fullName = String(payload.fullName || '').trim();
    const phone = String(payload.phone || '').trim();
    const role = String(payload.role || '');

    if (!email || !password || !fullName || !role) {
      return { success: false, message: 'Missing required fields.' };
    }
    if (role !== 'owner' && role !== 'barista') {
      return { success: false, message: 'Unknown role.' };
    }
    if (!phone) {
      return { success: false, message: 'A phone number is required.' };
    }
    if (password.length < 8) {
      return { success: false, message: 'Password should be at least 8 characters.' };
    }
    const cafeName = String(payload.cafeName || '').trim();
    if (role === 'owner' && !cafeName) {
      return { success: false, message: 'Café name is required for café owners.' };
    }
    const availabilityType = role === 'barista' ? String(payload.availabilityType || '').trim() : '';
    if (availabilityType && availabilityType !== 'part_time' && availabilityType !== 'full_time') {
      return { success: false, message: 'Unknown availability type.' };
    }

    const { data, error } = await db.auth.admin.createUser({
      email,
      password,
      email_confirm: true, // no email-verification flow in this app, same as the old prototype
      user_metadata: {
        role,
        full_name: fullName,
        phone,
        cafe_name: role === 'owner' ? cafeName : null
      }
    });

    if (error) {
      const alreadyRegistered = /already registered|already exists|email_exists/i.test(error.message || '');
      return {
        success: false,
        message: alreadyRegistered ? 'That email is already registered.' : 'Server error: ' + error.message
      };
    }

    // The on_auth_user_created trigger inserts the matching profiles row
    // (with role/full_name/phone/cafe_name from user_metadata above)
    // synchronously, so it's already there to read back.
    const { data: profile, error: profileErr } = await db
      .from('profiles')
      .select('*')
      .eq('id', data.user.id)
      .single();
    if (profileErr) throw profileErr;

    // availability_type isn't part of the signup trigger's metadata copy,
    // so it's set with a follow-up update rather than touching that trigger.
    let finalProfile = profile;
    if (availabilityType) {
      const { data: updated, error: updateErr } = await db
        .from('profiles')
        .update({ availability_type: availabilityType })
        .eq('id', data.user.id)
        .select()
        .single();
      if (updateErr) throw updateErr;
      finalProfile = updated;
    }

    return { success: true, user: profileToSafeUser(finalProfile) };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

async function loginUser(payload) {
  try {
    const db = getAdminClient();
    const email = String(payload.email || '').trim().toLowerCase();
    const password = String(payload.password || '');

    const { data, error } = await db.auth.signInWithPassword({ email, password });
    if (error || !data.user) {
      return { success: false, message: 'Email or password is incorrect.' };
    }

    const { data: profile, error: profileErr } = await db
      .from('profiles')
      .select('*')
      .eq('id', data.user.id)
      .single();
    if (profileErr) throw profileErr;

    return { success: true, user: profileToSafeUser(profile) };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

async function updateProfile(payload) {
  try {
    const db = getAdminClient();
    if (!payload.userId) return { success: false, message: 'User not found.' };

    const fullName = String(payload.fullName || '').trim();
    const phone = String(payload.phone || '').trim();
    if (!fullName) return { success: false, message: 'Full name is required.' };
    if (!phone) return { success: false, message: 'A phone number is required.' };

    const { data: existing, error: existingErr } = await db
      .from('profiles')
      .select('role, address')
      .eq('id', payload.userId)
      .maybeSingle();
    if (existingErr) throw existingErr;
    if (!existing) return { success: false, message: 'User not found.' };

    const role = payload.role || existing.role;
    const updates = { full_name: fullName, phone };

    if (role === 'owner') {
      const cafeName = String(payload.cafeName || '').trim();
      if (!cafeName) return { success: false, message: 'Café name is required.' };
      updates.cafe_name = cafeName;
      updates.bio = payload.bio || '';
      updates.logo_url = payload.logoUrl || '';
      updates.instagram_url = payload.instagramUrl || '';
      updates.facebook_url = payload.facebookUrl || '';
      updates.website_url = payload.websiteUrl || '';
      updates.cafe_images = payload.cafeImages || [];
      updates.address = payload.address || '';
    } else if (role === 'barista') {
      if (payload.availabilityType !== undefined) {
        const availabilityType = String(payload.availabilityType || '').trim();
        if (availabilityType && availabilityType !== 'part_time' && availabilityType !== 'full_time') {
          return { success: false, message: 'Unknown availability type.' };
        }
        updates.availability_type = availabilityType || null;
      }
      // A barista's address is their own base for distance estimates —
      // same column as a café's, but never shown to owners (see
      // listBaristas, which strips address/lat/lng before returning).
      updates.address = payload.address || '';
    }

    // Shared by both roles: resolve `updates.address` to lat/lng. Applies
    // whenever either role sets an address field above.
    if (updates.address !== undefined) {
      const addressChanged = updates.address !== (existing.address || '');
      const lat = Number(payload.lat);
      const lng = Number(payload.lng);
      const hasCoordsOverride = payload.lat != null && payload.lng != null && Number.isFinite(lat) && Number.isFinite(lng);

      if (hasCoordsOverride) {
        // Browser geolocation was used for this pin — trust it directly
        // rather than geocoding the typed address at all.
        updates.lat = lat;
        updates.lng = lng;
      } else if (addressChanged) {
        // Resolve the new address right away, synchronously, so whoever's
        // saving finds out immediately (via locationResolved on the
        // returned profile) if it didn't match anything — rather than
        // silently caching null and letting every distance estimate that
        // depends on it go quietly missing. getTravelEstimates still
        // re-attempts this lazily as a fallback for profiles saved before
        // this existed.
        updates.lat = null;
        updates.lng = null;
        if (updates.address) {
          try {
            const geocoded = await geocodeAddress(updates.address);
            if (geocoded) {
              updates.lat = geocoded.lat;
              updates.lng = geocoded.lng;
            }
          } catch (err) {
            // Save still proceeds with lat/lng left null — locationResolved
            // on the response surfaces this instead.
          }
        }
      }
    }

    const { data: updated, error: updateErr } = await db
      .from('profiles')
      .update(updates)
      .eq('id', payload.userId)
      .select()
      .single();
    if (updateErr) throw updateErr;

    return { success: true, user: profileToSafeUser(updated) };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

async function listBaristas() {
  try {
    const db = getAdminClient();
    const { data, error } = await db
      .from('profiles')
      .select('*')
      .eq('role', 'barista')
      .order('full_name', { ascending: true });
    if (error) throw error;

    // availabilityType is internal-only, and address is a barista's own
    // base for distance estimates — neither is ever shown to owners.
    const baristas = (data || []).map(profileToSafeUser).map(({ availabilityType, locationResolved, address, ...rest }) => rest);
    return { success: true, baristas };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * payload: { baristaUserId, ownerUserId? }
 * Powers the "worked N shifts across M cafés" line an owner sees when
 * considering a barista — derived entirely from completed shifts, no
 * stored counters to keep in sync.
 */
async function getBaristaStats(payload) {
  try {
    if (!payload || !payload.baristaUserId) {
      return { success: false, message: 'Barista User ID is missing.' };
    }

    const db = getAdminClient();
    const { data, error } = await db
      .from('shifts')
      .select('owner_id')
      .eq('assigned_barista_id', payload.baristaUserId)
      .eq('status', 'completed');
    if (error) throw error;

    const rows = data || [];
    const distinctCafes = new Set(rows.map(r => r.owner_id)).size;
    const shiftsAtThisCafe = payload.ownerUserId
      ? rows.filter(r => r.owner_id === payload.ownerUserId).length
      : null;

    return {
      success: true,
      stats: { totalCompletedShifts: rows.length, distinctCafes, shiftsAtThisCafe }
    };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * payload: { lat, lng, ownerUserIds }
 * Estimated driving time from a barista's current position (browser
 * geolocation — never stored) to each café owner's address. A café's
 * coordinates are geocoded once and cached on profiles.lat/lng
 * (updateProfile clears that cache when the address changes), so this
 * only calls out to LocationIQ for addresses it hasn't resolved yet.
 * Geocoding/routing failures are swallowed per-owner rather than failing
 * the whole batch — a missing ETA just means that shift shows no badge.
 */
async function getTravelEstimates(payload) {
  try {
    let lat = Number(payload && payload.lat);
    let lng = Number(payload && payload.lng);
    const hasLiveCoords = Number.isFinite(lat) && Number.isFinite(lng);
    const baristaUserId = payload && payload.baristaUserId;
    const ownerUserIds = Array.isArray(payload && payload.ownerUserIds)
      ? [...new Set(payload.ownerUserIds.filter(Boolean))]
      : [];
    if (!ownerUserIds.length) {
      return { success: true, estimates: {} };
    }

    const db = getAdminClient();

    // No live geolocation this time — fall back to wherever the barista
    // last saved their own address (profile screen, same lat/lng columns
    // as a café). If they've never set one either, there's simply nothing
    // to estimate from; that's reported back via baristaLocationMissing
    // rather than an error, so the UI can point them at their profile
    // instead of showing a broken feature.
    if (!hasLiveCoords) {
      if (!baristaUserId) {
        return { success: false, message: 'A current location is required.' };
      }
      const { data: self, error: selfErr } = await db
        .from('profiles')
        .select('lat, lng')
        .eq('id', baristaUserId)
        .maybeSingle();
      if (selfErr) throw selfErr;
      if (!self || self.lat == null || self.lng == null) {
        return { success: true, estimates: {}, baristaLocationMissing: true };
      }
      lat = self.lat;
      lng = self.lng;
    }

    const { data: owners, error } = await db
      .from('profiles')
      .select('id, address, lat, lng')
      .in('id', ownerUserIds);
    if (error) throw error;

    for (const owner of owners || []) {
      if (owner.lat != null && owner.lng != null) continue;
      if (!owner.address) continue;
      try {
        const geocoded = await geocodeAddress(owner.address);
        if (geocoded) {
          owner.lat = geocoded.lat;
          owner.lng = geocoded.lng;
          await db.from('profiles').update({ lat: geocoded.lat, lng: geocoded.lng }).eq('id', owner.id);
        }
      } catch (err) {
        // best-effort — leave this owner out of the matrix call below
      }
    }

    const located = (owners || []).filter(o => o.lat != null && o.lng != null);
    if (!located.length) {
      return { success: true, estimates: {} };
    }

    const estimates = {};
    try {
      const travelTimes = await getTravelTimes(
        { lat, lng },
        located.map(o => ({ lat: o.lat, lng: o.lng }))
      );
      located.forEach((owner, i) => {
        if (travelTimes[i]) estimates[owner.id] = travelTimes[i];
      });
    } catch (err) {
      // Matrix call failed outright (e.g. rate limit) — return no
      // estimates rather than erroring the whole shift list out.
    }

    return { success: true, estimates };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/* ================================================================
   SHIFTS — post, edit, list, status changes, rate bargaining
   ================================================================ */

async function getOwnerDashboard(payload) {
  try {
    if (typeof payload === 'string') payload = { ownerUserId: payload };
    if (!payload || !payload.ownerUserId) {
      return { success: false, message: 'Owner User ID is missing.' };
    }

    const db = getAdminClient();
    const { data, error } = await db
      .from('shift_details')
      .select('*')
      .eq('owner_id', payload.ownerUserId);
    if (error) throw error;

    const shifts = (data || []).map(shiftDetailsToObject);
    // 'requested' (awaiting the barista's response) stays with the open
    // shifts, not history — it's still an active shift the owner needs to
    // track, just not advertised to other baristas any more.
    const open = shifts.filter(s => s.status === 'open' || s.status === 'requested').sort((a, b) => (a.date < b.date ? -1 : 1));
    const history = shifts.filter(s => s.status !== 'open' && s.status !== 'requested').sort((a, b) => (a.date < b.date ? 1 : -1));

    return { success: true, openShifts: open, historyShifts: history };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

async function createShift(payload) {
  try {
    if (!payload.date || !payload.startTime || !payload.endTime) {
      return { success: false, message: 'Date, start time, and end time are required.' };
    }
    if (!payload.rate || Number(payload.rate) <= 0) {
      return { success: false, message: 'Enter a valid hourly rate.' };
    }
    if (Number(payload.rate) < MINIMUM_WAGE) {
      return { success: false, message: `Rate can't be below the minimum wage (£${MINIMUM_WAGE.toFixed(2)}/hr).` };
    }
    if (!payload.skills || !payload.skills.length) {
      return { success: false, message: 'Select at least one skill needed for this shift.' };
    }

    const db = getAdminClient();
    const { data, error } = await db
      .from('shifts')
      .insert({
        owner_id: payload.ownerUserId,
        cafe_name: payload.cafeName || '',
        shift_date: payload.date,
        start_time: payload.startTime,
        end_time: payload.endTime,
        rate: payload.rate,
        skills: payload.skills || [],
        busy_level: payload.busyLevel || 3,
        lunch_included: !!payload.lunchIncluded,
        break_minutes: payload.breakMinutes || 0,
        break_notes: payload.breakNotes || '',
        notes: payload.notes || '',
        rate_negotiable: payload.rateNegotiable !== false,
        status: 'open'
      })
      .select()
      .single();
    if (error) throw error;

    return { success: true, shift: shiftDetailsToObject(data) };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

async function updateShift(payload) {
  try {
    if (payload.rate !== undefined && Number(payload.rate) < MINIMUM_WAGE) {
      return { success: false, message: `Rate can't be below the minimum wage (£${MINIMUM_WAGE.toFixed(2)}/hr).` };
    }

    const db = getAdminClient();
    const { data, error } = await db
      .from('shifts')
      .update({
        shift_date: payload.date,
        start_time: payload.startTime,
        end_time: payload.endTime,
        rate: payload.rate,
        skills: payload.skills || [],
        busy_level: payload.busyLevel || 3,
        lunch_included: !!payload.lunchIncluded,
        break_minutes: payload.breakMinutes || 0,
        break_notes: payload.breakNotes || '',
        notes: payload.notes || '',
        rate_negotiable: payload.rateNegotiable !== false
      })
      .eq('id', payload.shiftId)
      .select()
      .maybeSingle();
    if (error) throw error;
    if (!data) return { success: false, message: 'Shift not found.' };

    const shift = await fetchShiftDetails_(payload.shiftId);
    return { success: true, shift };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * payload: { shiftId, status, assignedBaristaUserId? }
 * Generic status setter — used for marking a shift completed or cancelled.
 * Owner-initiated booking no longer goes through here directly (see
 * requestBooking/respondToBookingRequest below): a barista has to accept
 * before a shift the owner picked them for actually becomes 'filled'.
 * assignedBaristaName is accepted for backward compatibility but ignored,
 * since names are now joined live from profiles via shift_details, not
 * stored on the shift.
 */
async function setShiftStatus(payload) {
  try {
    const db = getAdminClient();
    const updates = { status: payload.status };
    if (payload.assignedBaristaUserId !== undefined) {
      updates.assigned_barista_id = payload.assignedBaristaUserId;
    }
    if (payload.status === 'cancelled') {
      updates.cancelled_by = 'owner';
    }

    const { data, error } = await db
      .from('shifts')
      .update(updates)
      .eq('id', payload.shiftId)
      .select()
      .maybeSingle();
    if (error) throw error;
    if (!data) return { success: false, message: 'Shift not found.' };

    const shift = await fetchShiftDetails_(payload.shiftId);
    return { success: true, shift };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

async function cancelShiftAsBarista(payload) {
  try {
    const db = getAdminClient();
    const current = await fetchShiftDetails_(payload.shiftId);
    if (!current) return { success: false, message: 'This shift no longer exists.' };
    if (current.status !== 'filled') {
      return { success: false, message: 'This shift is not currently assigned to you.' };
    }
    if (String(current.assignedBaristaUserId) !== String(payload.baristaUserId)) {
      return { success: false, message: 'This shift is not assigned to you.' };
    }

    const { error } = await db
      .from('shifts')
      .update({ status: 'cancelled', cancelled_by: 'barista' })
      .eq('id', payload.shiftId);
    if (error) throw error;

    const shift = await fetchShiftDetails_(payload.shiftId);
    return { success: true, shift };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

async function getBaristaShifts(payload) {
  try {
    if (typeof payload === 'string') payload = { baristaUserId: payload };
    if (!payload || !payload.baristaUserId) {
      return { success: false, message: 'Barista User ID is missing.' };
    }

    const db = getAdminClient();
    const { data, error } = await db
      .from('shift_details')
      .select('*')
      .eq('assigned_barista_id', payload.baristaUserId);
    if (error) throw error;

    const shifts = (data || []).map(shiftDetailsToObject);
    const requested = shifts.filter(s => s.status === 'requested');
    const filled = shifts.filter(s => s.status === 'filled');
    const completed = shifts.filter(s => s.status === 'completed' && s.baristaRatingOfCafe === null);

    requested.sort((a, b) => (a.date === b.date ? (a.startTime < b.startTime ? -1 : 1) : (a.date < b.date ? -1 : 1)));
    filled.sort((a, b) => (a.date === b.date ? (a.startTime < b.startTime ? -1 : 1) : (a.date < b.date ? -1 : 1)));
    completed.sort((a, b) => (a.date === b.date ? 0 : (a.date < b.date ? 1 : -1)));

    return { success: true, shifts: requested.concat(filled, completed) };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

async function repostShift(payload) {
  try {
    const db = getAdminClient();
    const original = await fetchShiftDetails_(payload.shiftId);
    if (!original) return { success: false, message: 'Shift not found.' };
    if (original.status !== 'cancelled' || original.cancelledBy !== 'barista') {
      return { success: false, message: 'Only a barista-cancelled shift can be reposted.' };
    }
    if (original.repostedShiftId) {
      return { success: false, message: 'This shift has already been reposted.' };
    }

    const { data: created, error: createErr } = await db
      .from('shifts')
      .insert({
        owner_id: original.ownerUserId,
        cafe_name: original.cafeName,
        shift_date: original.date,
        start_time: original.startTime,
        end_time: original.endTime,
        rate: original.rate,
        skills: original.skills || [],
        busy_level: original.busyLevel,
        lunch_included: !!original.lunchIncluded,
        break_minutes: original.breakMinutes,
        break_notes: original.breakNotes,
        notes: original.notes,
        status: 'open'
      })
      .select()
      .single();
    if (createErr) throw createErr;

    const { error: linkErr } = await db
      .from('shifts')
      .update({ reposted_shift_id: created.id })
      .eq('id', payload.shiftId);
    if (linkErr) throw linkErr;

    const shift = await fetchShiftDetails_(created.id);
    return { success: true, shift };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

async function dismissRepostPrompt(payload) {
  try {
    const db = getAdminClient();
    // No free-text sentinel column to lean on any more (RepostedShiftID
    // used to double as a "dismissed" flag) — a real boolean column that
    // means "owner has acted on this cancellation" would be cleaner, but
    // that's a schema change outside this refactor's scope. For now this
    // links the cancelled shift to itself, which reads the same way
    // repostShift's check does (`repostedShiftId` truthy) without needing
    // a new column.
    const { data, error } = await db
      .from('shifts')
      .update({ reposted_shift_id: payload.shiftId })
      .eq('id', payload.shiftId)
      .select()
      .maybeSingle();
    if (error) throw error;
    if (!data) return { success: false, message: 'Shift not found.' };

    const shift = await fetchShiftDetails_(payload.shiftId);
    return { success: true, shift };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

async function rateShift(payload) {
  try {
    const db = getAdminClient();
    const shift = await fetchShiftDetails_(payload.shiftId);
    if (!shift) return { success: false, message: 'Shift not found.' };

    const rating = Math.round(Number(payload.rating));
    if (!rating || rating < 1 || rating > 5) {
      return { success: false, message: 'Rating must be between 1 and 5 stars.' };
    }
    if (shift.status !== 'completed') {
      return { success: false, message: 'Only completed shifts can be rated.' };
    }

    let updates;
    let targetUserId;
    if (payload.raterRole === 'barista') {
      if (String(shift.assignedBaristaUserId) !== String(payload.raterUserId)) {
        return { success: false, message: 'This shift is not assigned to you.' };
      }
      if (shift.baristaRatingOfCafe !== null) {
        return { success: false, message: "You've already rated this shift." };
      }
      updates = { barista_rating_of_cafe: rating };
      targetUserId = shift.ownerUserId;
    } else if (payload.raterRole === 'owner') {
      if (String(shift.ownerUserId) !== String(payload.raterUserId)) {
        return { success: false, message: "This isn't your shift to rate." };
      }
      if (shift.ownerRatingOfBarista !== null) {
        return { success: false, message: "You've already rated this shift." };
      }
      if (!shift.assignedBaristaUserId) {
        return { success: false, message: 'This shift has no assigned barista to rate.' };
      }
      updates = { owner_rating_of_barista: rating };
      targetUserId = shift.assignedBaristaUserId;
    } else {
      return { success: false, message: 'Unknown rater role.' };
    }

    const { error } = await db.from('shifts').update(updates).eq('id', payload.shiftId);
    if (error) throw error;

    await incrementUserRating_(targetUserId, rating);

    const updated = await fetchShiftDetails_(payload.shiftId);
    return { success: true, shift: updated };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

// Not atomic (read-then-write), same as the old Apps Script version — a
// real fix would be a Postgres RPC doing the increment in one statement,
// but that's a schema change outside this refactor's scope.
async function incrementUserRating_(userId, rating) {
  const db = getAdminClient();
  const { data, error } = await db
    .from('profiles')
    .select('rating_total, rating_count')
    .eq('id', userId)
    .maybeSingle();
  if (error) throw error;
  if (!data) return;

  const newTotal = (Number(data.rating_total) || 0) + rating;
  const newCount = (Number(data.rating_count) || 0) + 1;
  const { error: updateErr } = await db
    .from('profiles')
    .update({ rating_total: newTotal, rating_count: newCount })
    .eq('id', userId);
  if (updateErr) throw updateErr;
}

async function proposeRate(payload) {
  try {
    const db = getAdminClient();
    const current = await fetchShiftDetails_(payload.shiftId);
    if (!current) return { success: false, message: 'This shift no longer exists.' };
    if (current.status !== 'open') {
      return { success: false, message: 'This shift is no longer open.' };
    }
    if (current.rateNegotiable === false) {
      return { success: false, message: "This shift's rate is fixed and not open to offers." };
    }

    const proposedRate = Number(payload.proposedRate);
    if (!proposedRate || proposedRate <= 0) {
      return { success: false, message: 'Enter a valid hourly rate.' };
    }
    if (proposedRate < MINIMUM_WAGE) {
      return { success: false, message: `Your rate can't be below the minimum wage (£${MINIMUM_WAGE.toFixed(2)}/hr).` };
    }

    const { error } = await db
      .from('shifts')
      .update({
        proposed_rate: proposedRate,
        proposed_by_id: payload.baristaUserId || null,
        negotiation_status: 'pending'
      })
      .eq('id', payload.shiftId);
    if (error) throw error;

    const shift = await fetchShiftDetails_(payload.shiftId);
    return { success: true, shift };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

async function respondToRateProposal(payload) {
  try {
    const db = getAdminClient();
    const current = await fetchShiftDetails_(payload.shiftId);
    if (!current) return { success: false, message: 'Shift not found.' };
    if (current.negotiationStatus !== 'pending') {
      return { success: false, message: 'There is no pending rate offer on this shift.' };
    }

    let updates;
    if (payload.accept) {
      updates = {
        rate: current.proposedRate,
        status: 'filled',
        assigned_barista_id: current.proposedByUserId,
        negotiation_status: 'accepted'
      };
    } else {
      updates = {
        negotiation_status: 'declined',
        proposed_rate: null,
        proposed_by_id: null
      };
    }

    const { error } = await db.from('shifts').update(updates).eq('id', payload.shiftId);
    if (error) throw error;

    const shift = await fetchShiftDetails_(payload.shiftId);
    return { success: true, shift };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

async function listOpenShifts() {
  try {
    const db = getAdminClient();
    const today = new Date().toISOString().slice(0, 10);

    const { data, error } = await db
      .from('shift_details')
      .select('*')
      .eq('status', 'open')
      .gte('shift_date', today);
    if (error) throw error;

    const open = (data || []).map(shiftDetailsToObject);
    open.sort((a, b) => (a.date === b.date ? (a.startTime < b.startTime ? -1 : 1) : (a.date < b.date ? -1 : 1)));

    return { success: true, shifts: open, minimumWage: MINIMUM_WAGE };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

async function acceptShift(payload) {
  try {
    const db = getAdminClient();
    const current = await fetchShiftDetails_(payload.shiftId);
    if (!current) return { success: false, message: 'This shift no longer exists.' };
    if (current.status !== 'open') {
      return { success: false, message: 'This shift was just taken by someone else.' };
    }

    const { error } = await db
      .from('shifts')
      .update({ status: 'filled', assigned_barista_id: payload.baristaUserId || null })
      .eq('id', payload.shiftId);
    if (error) throw error;

    const shift = await fetchShiftDetails_(payload.shiftId);
    return { success: true, shift };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * payload: { shiftId, baristaUserId }
 * Owner picks a barista for an open shift. This no longer books them
 * outright — it moves the shift to 'requested' and waits for the barista
 * to accept via respondToBookingRequest, so a barista can no longer be
 * assigned to a shift without their say-so.
 */
async function requestBooking(payload) {
  try {
    if (!payload || !payload.baristaUserId) {
      return { success: false, message: 'Barista is missing.' };
    }

    const db = getAdminClient();
    const current = await fetchShiftDetails_(payload.shiftId);
    if (!current) return { success: false, message: 'This shift no longer exists.' };
    if (current.status !== 'open') {
      return { success: false, message: 'This shift is no longer open.' };
    }

    const { error } = await db
      .from('shifts')
      .update({ status: 'requested', assigned_barista_id: payload.baristaUserId })
      .eq('id', payload.shiftId);
    if (error) throw error;

    const shift = await fetchShiftDetails_(payload.shiftId);
    return { success: true, shift };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * payload: { shiftId, baristaUserId, accept }
 * The barista's response to a requestBooking — only they can accept or
 * decline a request addressed to them.
 */
async function respondToBookingRequest(payload) {
  try {
    const db = getAdminClient();
    const current = await fetchShiftDetails_(payload.shiftId);
    if (!current) return { success: false, message: 'This shift no longer exists.' };
    if (current.status !== 'requested') {
      return { success: false, message: 'This request is no longer pending.' };
    }
    if (String(current.assignedBaristaUserId) !== String(payload.baristaUserId)) {
      return { success: false, message: 'This request is not addressed to you.' };
    }

    const updates = payload.accept
      ? { status: 'filled' }
      : { status: 'open', assigned_barista_id: null };

    const { error } = await db.from('shifts').update(updates).eq('id', payload.shiftId);
    if (error) throw error;

    const shift = await fetchShiftDetails_(payload.shiftId);
    return { success: true, shift };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * payload: { shiftId }
 * Owner-side cancel of a still-pending request, reopening the shift to
 * other baristas rather than cancelling the whole shift.
 */
async function withdrawBookingRequest(payload) {
  try {
    const db = getAdminClient();
    const current = await fetchShiftDetails_(payload.shiftId);
    if (!current) return { success: false, message: 'This shift no longer exists.' };
    if (current.status !== 'requested') {
      return { success: false, message: 'This shift has no pending request.' };
    }

    const { error } = await db
      .from('shifts')
      .update({ status: 'open', assigned_barista_id: null })
      .eq('id', payload.shiftId);
    if (error) throw error;

    const shift = await fetchShiftDetails_(payload.shiftId);
    return { success: true, shift };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/* ================================================================
   IMAGE UPLOAD — café logo + photos, stored in Supabase Storage
   ================================================================ */

async function uploadImage(payload) {
  try {
    if (!payload.base64 || !payload.mimeType) {
      return { success: false, message: 'No image data received.' };
    }

    const buffer = Buffer.from(payload.base64, 'base64');
    if (buffer.length > MAX_IMAGE_SIZE_BYTES) {
      const maxMb = (MAX_IMAGE_SIZE_BYTES / (1024 * 1024)).toFixed(0);
      return { success: false, message: `Image is too large — please use a file under ${maxMb}MB.` };
    }

    const ext = (payload.fileName && payload.fileName.includes('.'))
      ? payload.fileName.split('.').pop()
      : (payload.mimeType.split('/')[1] || 'jpg');
    const path = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}.${ext}`;

    const db = getAdminClient();
    const { error: uploadErr } = await db.storage
      .from(STORAGE_BUCKET)
      .upload(path, buffer, { contentType: payload.mimeType, upsert: false });
    if (uploadErr) throw uploadErr;

    const { data } = db.storage.from(STORAGE_BUCKET).getPublicUrl(path);
    return { success: true, url: data.publicUrl };
  } catch (err) {
    return { success: false, message: 'Upload failed: ' + err.message };
  }
}

module.exports = {
  registerUser,
  loginUser,
  updateProfile,
  listBaristas,
  getBaristaStats,
  getTravelEstimates,
  getOwnerDashboard,
  createShift,
  updateShift,
  setShiftStatus,
  requestBooking,
  respondToBookingRequest,
  withdrawBookingRequest,
  cancelShiftAsBarista,
  getBaristaShifts,
  repostShift,
  dismissRepostPrompt,
  rateShift,
  proposeRate,
  respondToRateProposal,
  listOpenShifts,
  acceptShift,
  uploadImage
};
