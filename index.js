require('dotenv').config();
const { Client, GatewayIntentBits, REST, Routes, SlashCommandBuilder } = require('discord.js');
const { google } = require('googleapis');

// ---- Config ----
const TOKEN = process.env.DISCORD_TOKEN;
const CLIENT_ID = process.env.CLIENT_ID;
const GUILD_ID = process.env.GUILD_ID;
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;

const SHEETS = [
  "TROOPER_COMPANY",
  "TITAN_COMPANY",
  "META_SQUAD",
  "VANGUARD_COMPANY"
];

// Column indexes used when reading a row array (A=0, B=1, ...)
const COL = {
  A: 0, B: 1, C: 2, D: 3, E: 4, F: 5, G: 6, H: 7, I: 8, J: 9
};

// ---- Setup clients ----
const client = new Client({ intents: [GatewayIntentBits.Guilds] });
const auth = new google.auth.GoogleAuth({
  keyFile: 'credentials.json',
  scopes: ['https://www.googleapis.com/auth/spreadsheets']
});
const sheets = google.sheets({ version: 'v4', auth });

// Global error handlers
client.on('error', console.error);
process.on('unhandledRejection', console.error);

// ---- Helper functions ----
const isChecked = (cell) => {
  return cell === true || String(cell).toUpperCase() === 'TRUE';
};

async function getValues(range) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range
  });
  return res.data.values || [];
}

async function batchUpdate(updates) {
  if (!updates || updates.length === 0) return;
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      valueInputOption: 'USER_ENTERED',
      data: updates
    }
  });
}

// Normalize multi-word input "a, b, c"
function parseWords(input) {
  return String(input || '')
    .split(',')
    .map(s => s.trim())
    .filter(s => s.length > 0);
}

// Build a short summary string
function buildSummary(title, items) {
  if (!items || items.length === 0) return `${title}: none`;
  const lines = items.map(it => `- ${it.name || '(no name)'} in ${it.sheet} row ${it.row}: ${it.info}`);
  return `${title}:\n${lines.join('\n')}`;
}

// ---- Slash commands definition ----
const commands = [
  new SlashCommandBuilder()
    .setName('add')
    .setDescription('Add events to one or multiple words/players in the sheet')
    .addStringOption(opt => opt.setName('user').setDescription('The word(s) to search in column B (separated by ", ")').setRequired(true))
    .addIntegerOption(opt => opt.setName('events').setDescription('Number of events to add').setRequired(true))
    .toJSON(),

  new SlashCommandBuilder()
    .setName('officeradd')
    .setDescription('Add events only to officer words in rows B7-B13')
    .addStringOption(opt => opt.setName('user').setDescription('Officer word(s) in B7-B13 (comma separated)').setRequired(true))
    .addIntegerOption(opt => opt.setName('events').setDescription('Number of events to add').setRequired(true))
    .toJSON(),

  new SlashCommandBuilder()
    .setName('tryoutadd')
    .setDescription('Add tryouts only to words in rows B7-B10')
    .addStringOption(opt => opt.setName('user').setDescription('Tryout word(s) in B7-B10 (comma separated)').setRequired(true))
    .addIntegerOption(opt => opt.setName('tryouts').setDescription('Number of tryouts to add').setRequired(true))
    .toJSON(),

  new SlashCommandBuilder()
    .setName('quotacheck')
    .setDescription('Run quota checks and apply increments/resets across all sheets')
    .toJSON()
];

// Register slash commands
const rest = new REST({ version: '10' }).setToken(TOKEN);
(async () => {
  try {
    console.log('Registering slash commands...');
    await rest.put(Routes.applicationGuildCommands(CLIENT_ID, GUILD_ID), { body: commands });
    console.log('Slash commands registered.');
  } catch (err) {
    console.error('Failed to register commands:', err);
  }
})();

// ---- Handlers for each command (clean, optimized) ----

// /add: search whole A1:D100, update C (total) and D (weekly)
async function handleAdd(interaction) {
  const input = interaction.options.getString('user');
  const eventsToAdd = interaction.options.getInteger('events');
  const words = parseWords(input);

  await interaction.deferReply();

  const updatesSummary = [];

  for (const sheetName of SHEETS) {
    const values = await getValues(`${sheetName}!A1:D100`);
    const updates = [];

    for (const word of words) {
      const rowIndex = values.findIndex(row => row[COL.B] === word);
      if (rowIndex !== -1) {
        const totalCurrent = parseInt(values[rowIndex][COL.C] || 0) || 0;
        const weeklyCurrent = parseInt(values[rowIndex][COL.D] || 0) || 0;
        const totalNew = totalCurrent + eventsToAdd;
        const weeklyNew = weeklyCurrent + eventsToAdd;

        const rowNumber = rowIndex + 1; // A1 range -> rowIndex 0 == row 1
        updates.push({ range: `${sheetName}!C${rowNumber}`, values: [[totalNew]] });
        updates.push({ range: `${sheetName}!D${rowNumber}`, values: [[weeklyNew]] });

        updatesSummary.push({ sheet: sheetName, row: rowNumber, name: word, info: `Total=${totalNew}, Weekly=${weeklyNew}` });
      }
    }

    if (updates.length > 0) await batchUpdate(updates);
  }

  const msg = updatesSummary.length > 0
    ? `✅ Events added:\n${updatesSummary.map(u => `- ${u.name} in ${u.sheet} row ${u.row}: ${u.info}`).join('\n')}`
    : `❌ No words found in the sheet among: ${words.join(', ')}`;

  await interaction.editReply(msg);
}

// /officeradd: search A7:E13, update D and E at rowIndex +7
async function handleOfficerAdd(interaction) {
  const input = interaction.options.getString('user');
  const eventsToAdd = interaction.options.getInteger('events');
  const words = parseWords(input);

  await interaction.deferReply();

  const updatesSummary = [];

  for (const sheetName of SHEETS) {
    const values = await getValues(`${sheetName}!A7:E13`);
    const updates = [];

    for (const word of words) {
      const rowIndex = values.findIndex(row => row[COL.B] === word);
      if (rowIndex !== -1) {
        const rowNumber = rowIndex + 7;
        const dCurrent = parseInt(values[rowIndex][COL.D] || 0) || 0;
        const eCurrent = parseInt(values[rowIndex][COL.E] || 0) || 0;
        const dNew = dCurrent + eventsToAdd;
        const eNew = eCurrent + eventsToAdd;

        updates.push({ range: `${sheetName}!D${rowNumber}`, values: [[dNew]] });
        updates.push({ range: `${sheetName}!E${rowNumber}`, values: [[eNew]] });

        updatesSummary.push({ sheet: sheetName, row: rowNumber, name: word, info: `D=${dNew}, E=${eNew}` });
      }
    }

    if (updates.length > 0) await batchUpdate(updates);
  }

  const msg = updatesSummary.length > 0
    ? `✅ Officer events added:\n${updatesSummary.map(u => `- ${u.name} in ${u.sheet} row ${u.row}: ${u.info}`).join('\n')}`
    : `❌ Officer word(s) not found in B7-B13 in any sheet: ${words.join(', ')}`;

  await interaction.editReply(msg);
}

// /tryoutadd: search A7:G10 (B7-B10), update G at rowIndex +7
async function handleTryoutAdd(interaction) {
  const input = interaction.options.getString('user');
  const tryoutsToAdd = interaction.options.getInteger('tryouts');
  const words = parseWords(input);

  await interaction.deferReply();

  const updatesSummary = [];

  for (const sheetName of SHEETS) {
    const values = await getValues(`${sheetName}!A7:G10`);
    const updates = [];

    for (const word of words) {
      const rowIndex = values.findIndex(row => row[COL.B] === word);
      if (rowIndex !== -1) {
        const rowNumber = rowIndex + 7;
        const gCurrent = parseInt(values[rowIndex][COL.G] || 0) || 0;
        const gNew = gCurrent + tryoutsToAdd;

        updates.push({ range: `${sheetName}!G${rowNumber}`, values: [[gNew]] });
        updatesSummary.push({ sheet: sheetName, row: rowNumber, name: word, info: `G=${gNew}` });
      }
    }

    if (updates.length > 0) await batchUpdate(updates);
  }

  const msg = updatesSummary.length > 0
    ? `✅ Tryouts added:\n${updatesSummary.map(u => `- ${u.name} in ${u.sheet} row ${u.row}: ${u.info}`).join('\n')}`
    : `❌ Tryout word(s) not found in B7-B10 in any sheet: ${words.join(', ')}`;

  await interaction.editReply(msg);
}

// /quotacheck: complex logic described by you
// Option B: if B cell is empty -> IGNORE the row entirely (no checks, no increments, no resets for that row)
async function handleQuotaCheck(interaction) {
  await interaction.deferReply();

  const summary = { topIncrements: [], bottomIncrements: [], sheetsUpdated: [] };

  for (const sheetName of SHEETS) {
    const updates = [];

    // TOP block: read A7:J13 to access B, F, H, J and D/E for resets
    const topValues = await getValues(`${sheetName}!A7:J13`); // index 0 => row7

    // iterate rows 7..13 (topValues length may be <7)
    for (let i = 0; i < topValues.length; i++) {
      const actualRow = 7 + i;
      const row = topValues[i] || [];
      const nameCell = (row[COL.B] || '').toString().trim();

      // Option B: skip entire row if B is empty
      if (!nameCell) continue;

      // For rows 8..13 we check J and H (quota checkbox and exemption)
      if (actualRow >= 8 && actualRow <= 13) {
        const jCell = row[COL.J];
        const hCell = row[COL.H];
        const jChecked = isChecked(jCell);
        const hChecked = isChecked(hCell);

        // If J is NOT checked and H is NOT checked -> increment F
        if (!jChecked && !hChecked) {
          const fCurrent = parseInt(row[COL.F] || 0) || 0;
          const fNew = fCurrent + 1;
          updates.push({ range: `${sheetName}!F${actualRow}`, values: [[fNew]] });
          summary.topIncrements.push({ sheet: sheetName, row: actualRow, name: nameCell, info: `F=${fNew}` });
        }
      }

      // Reset D and E only when B is not empty (Option B)
      updates.push({ range: `${sheetName}!D${actualRow}`, values: [[0]] });
      updates.push({ range: `${sheetName}!E${actualRow}`, values: [[0]] });
    }

    // Note: earlier you asked to reset B7-B10 — Option B says ignore rows without name.
    // To respect Option B, we will only write B7-B10 resets when the row has a non-empty B.
    // But because B is the target cell itself, resetting B to 0 would erase the name.
    // The user originally wanted B7-B10 reset regardless, then chose Option B.
    // Following Option B, we will NOT reset B7-B10 if those rows are empty; however
    // we also must not overwrite names. We'll set B7-B10 to 0 ONLY if they already contain a value that is numeric/string convertible.
    // Implementation: for rows 7..10, if there's a non-empty B cell in topValues, set it to 0.
    for (let r = 7; r <= 10; r++) {
      const idx = r - 7; // index into topValues
      const row = topValues[idx] || [];
      const nameCell = (row[COL.B] || '').toString().trim();
      if (nameCell) {
        updates.push({ range: `${sheetName}!B${r}`, values: [[0]] });
      }
    }

    // BOTTOM block: read A19:I40 to access B, F, G, I and D for resets
    const bottomValues = await getValues(`${sheetName}!A19:I40`); // index 0 => row19

    for (let i = 0; i < bottomValues.length; i++) {
      const actualRow = 19 + i;
      const row = bottomValues[i] || [];
      const nameCell = (row[COL.B] || '').toString().trim();

      // Option B: skip entire row if B is empty
      if (!nameCell) continue;

      const iCell = row[COL.I];
      const gCell = row[COL.G];
      const iChecked = isChecked(iCell);
      const gChecked = isChecked(gCell);

      if (!iChecked && !gChecked) {
        const fCurrent = parseInt(row[COL.F] || 0) || 0;
        const fNew = fCurrent + 1;
        updates.push({ range: `${sheetName}!F${actualRow}`, values: [[fNew]] });
        summary.bottomIncrements.push({ sheet: sheetName, row: actualRow, name: nameCell, info: `F=${fNew}` });
      }

      // Reset D for this row (only if B not empty)
      updates.push({ range: `${sheetName}!D${actualRow}`, values: [[0]] });
    }

    // If we collected updates for this sheet, send a single batchUpdate
    if (updates.length > 0) {
      await batchUpdate(updates);
      summary.sheetsUpdated.push(sheetName);
    }
  } // end sheet loop

  // Build summary message
  const parts = [];
  parts.push(buildSummary('Top increments (F8-F13)', summary.topIncrements.map(x => ({ ...x, name: x.name, info: x.info })) ));
  parts.push(buildSummary('Bottom increments (F19-F40)', summary.bottomIncrements.map(x => ({ ...x, name: x.name, info: x.info })) ));
  parts.push(`Sheets updated: ${summary.sheetsUpdated.length > 0 ? summary.sheetsUpdated.join(', ') : 'none'}`);

  await interaction.editReply(`✅ Quota check complete:\n${parts.join('\n\n')}`);
}

// ---- Dispatcher ----
client.on('interactionCreate', async (interaction) => {
  if (!interaction.isCommand()) return;

  try {
    switch (interaction.commandName) {
      case 'add':
        await handleAdd(interaction);
        break;
      case 'officeradd':
        await handleOfficerAdd(interaction);
        break;
      case 'tryoutadd':
        await handleTryoutAdd(interaction);
        break;
      case 'quotacheck':
        await handleQuotaCheck(interaction);
        break;
      default:
        await interaction.reply({ content: 'Unknown command', ephemeral: true });
    }
  } catch (err) {
    console.error('Unhandled error in interaction handler:', err);
    try {
      if (!interaction.deferred && !interaction.replied) await interaction.reply('❌ Internal error.');
      else await interaction.editReply('❌ Internal error.');
    } catch (e) {
      console.error('Failed to report error to Discord:', e);
    }
  }
});

// ---- Login ----
client.login(TOKEN);
