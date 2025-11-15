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

// /quotacheck: corrected & optimized
async function handleQuotaCheck(interaction) {
  await interaction.deferReply();

  const summary = {
    topIncrements: [],
    bottomIncrements: [],
    sheetsUpdated: []
  };

  for (const sheetName of SHEETS) {
    const updates = [];

    // ---- TOP BLOCK (Rows 7–13) ----
    const topValues = await getValues(`${sheetName}!A7:J13`);

    for (let i = 0; i < topValues.length; i++) {
      const rowNumber = 7 + i;
      const row = topValues[i] || [];
      const username = (row[COL.B] || '').trim();

      // Skip empty rows entirely
      if (!username) continue;

      // Rows 8–13 quotas (F increments)
      if (rowNumber >= 8 && rowNumber <= 13) {
        const quotaPassed = isChecked(row[COL.J]);
        const exempt = isChecked(row[COL.H]);

        if (!quotaPassed && !exempt) {
          const fCurrent = Number(row[COL.F] || 0);
          const fNew = fCurrent + 1;

          updates.push({ range: `${sheetName}!F${rowNumber}`, values: [[fNew]] });
          summary.topIncrements.push({ sheet: sheetName, row: rowNumber, name: username, info: `F=${fNew}` });
        }
      }

      // Reset D & E (only if username exists)
      updates.push({ range: `${sheetName}!D${rowNumber}`, values: [[0]] });
      updates.push({ range: `${sheetName}!E${rowNumber}`, values: [[0]] });
    }

    // ---- Tryout reset: ONLY G7–G10 ----
    for (let r = 7; r <= 10; r++) {
      const row = topValues[r - 7] || [];
      const username = (row[COL.B] || '').trim();

      if (username) {
        updates.push({ range: `${sheetName}!G${r}`, values: [[0]] });
      }
    }

    // ---- BOTTOM BLOCK (19–40) ----
    const bottomValues = await getValues(`${sheetName}!A19:I40`);

    for (let i = 0; i < bottomValues.length; i++) {
      const rowNumber = 19 + i;
      const row = bottomValues[i] || [];
      const username = (row[COL.B] || '').trim();

      // Skip empty names
      if (!username) continue;

      // Quota logic for troopers
      const quotaPassed = isChecked(row[COL.I]);
      const exempt = isChecked(row[COL.G]);

      if (!quotaPassed && !exempt) {
        const fCurrent = Number(row[COL.F] || 0);
        const fNew = fCurrent + 1;

        updates.push({ range: `${sheetName}!F${rowNumber}`, values: [[fNew]] });
        summary.bottomIncrements.push({ sheet: sheetName, row: rowNumber, name: username, info: `F=${fNew}` });
      }

      // Reset D only if username exists
      updates.push({ range: `${sheetName}!D${rowNumber}`, values: [[0]] });
    }

    // Commit updates for this sheet
    if (updates.length > 0) {
      await batchUpdate(updates);
      summary.sheetsUpdated.push(sheetName);
    }
  }

  await interaction.editReply(
    `✅ **Quota check complete**\n\n` +
    `**Top increments:**\n${summary.topIncrements.map(x => `- ${x.name} (${x.sheet} R${x.row}) → ${x.info}`).join('\n') || 'none'}\n\n` +
    `**Bottom increments:**\n${summary.bottomIncrements.map(x => `- ${x.name} (${x.sheet} R${x.row}) → ${x.info}`).join('\n') || 'none'}\n\n` +
    `**Sheets updated:** ${summary.sheetsUpdated.join(', ') || 'none'}`
  );
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
