/* =====================================================================
   Work activities and safety documents.
   Edit this file to add, rename or remove forms — no other code changes needed.

   Activity:  { id, label }
   Document:  { id, title, desc,
                always:   true  -> required on every job
                when:     [activity ids] -> required when any of these is ticked
                optional: true  -> never required (e.g. incident report)
                page:     'forms/xyz.html' once the fillable form is built
                          (until then it opens a "coming soon" page) }
   ===================================================================== */

window.SAFETY_ACTIVITIES = [
  { id: 'climb',      label: 'Tower climbing' },
  { id: 'roof',       label: 'Rooftop work' },
  { id: 'rf',         label: 'Work near RF antennas' },
  { id: 'electrical', label: 'Electrical work' },
  { id: 'hotwork',    label: 'Hot work (welding, cutting, grinding)' },
  { id: 'lift',       label: 'Crane, hoisting or rigging' },
  { id: 'alone',      label: 'Working alone' }
];

window.SAFETY_DOCS = [
  { id: 'flha',           title: 'Field Level Hazard Assessment (FLHA)',  desc: 'Hazards and controls for today’s work', always: true },
  { id: 'toolbox',        title: 'Toolbox talk / tailgate meeting',       desc: 'Crew briefing and sign-in', always: true },
  { id: 'erp',            title: 'Site emergency response plan',          desc: 'Muster point and emergency contacts', always: true },
  { id: 'tower-access',   title: 'Pre-climb attestation (Tower Access)',  desc: 'Bell Mobility pre-climb health & safety attestation', when: ['climb'] },
  { id: 'fall-plan',      title: 'Fall protection plan',                  desc: 'Site- and structure-specific plan', when: ['climb', 'roof'] },
  { id: 'rescue-plan',    title: 'Rescue plan',                           desc: 'Rescuers, equipment and procedure', when: ['climb', 'roof'] },
  { id: 'structure',      title: 'Pre-climb structure inspection',        desc: 'Structure condition before climbing', when: ['climb'] },
  { id: 'ppe-inspection', title: 'Fall protection equipment inspection',  desc: 'Harness, lanyards, connectors', when: ['climb', 'roof'] },
  { id: 'rf',             title: 'RF safety checklist',                   desc: 'Exposure limits, power-down, monitors', when: ['rf'] },
  { id: 'electrical',     title: 'Electrical safety / lockout checklist', desc: 'Isolation and lockout verification', when: ['electrical'] },
  { id: 'hot-work',       title: 'Hot work permit',                       desc: 'Fire watch and permit sign-off', when: ['hotwork'] },
  { id: 'lift-plan',      title: 'Lift / rigging plan',                   desc: 'Loads, rigging and signals', when: ['lift'] },
  { id: 'working-alone',  title: 'Working alone check-in plan',           desc: 'Check-in times and contact', when: ['alone'] },
  { id: 'incident',       title: 'Incident / near-miss report',           desc: 'Only if something happens', optional: true },
  { id: 'visitor',        title: 'Site visitor sign-in',                  desc: 'Anyone visiting the site', optional: true }
];
