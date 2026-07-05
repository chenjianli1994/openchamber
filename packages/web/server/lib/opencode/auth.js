import fs from 'fs';
import { getLangCoderOpencodeAuthFilePath, getLangCoderOpencodeDataDir } from '../langcoder/opencode-runtime.js';

function getAuthFilePath() {
  return getLangCoderOpencodeAuthFilePath(process.env);
}

function getOpencodeDataDir() {
  return getLangCoderOpencodeDataDir(process.env);
}

function readAuthFile() {
  const authFile = getAuthFilePath();
  if (!fs.existsSync(authFile)) {
    return {};
  }
  try {
    const content = fs.readFileSync(authFile, 'utf8');
    const trimmed = content.trim();
    if (!trimmed) {
      return {};
    }
    return JSON.parse(trimmed);
  } catch (error) {
    console.error('Failed to read auth file:', error);
    throw new Error('Failed to read OpenCode auth configuration');
  }
}

function writeAuthFile(auth) {
  const opencodeDataDir = getOpencodeDataDir();
  const authFile = getAuthFilePath();
  try {
    if (!fs.existsSync(opencodeDataDir)) {
      fs.mkdirSync(opencodeDataDir, { recursive: true });
    }

    if (fs.existsSync(authFile)) {
      const backupFile = `${authFile}.openchamber.backup`;
      fs.copyFileSync(authFile, backupFile);
      console.log(`Created auth backup: ${backupFile}`);
    }

    fs.writeFileSync(authFile, JSON.stringify(auth, null, 2), 'utf8');
    console.log('Successfully wrote auth file');
  } catch (error) {
    console.error('Failed to write auth file:', error);
    throw new Error('Failed to write OpenCode auth configuration');
  }
}

function removeProviderAuth(providerId) {
  if (!providerId || typeof providerId !== 'string') {
    throw new Error('Provider ID is required');
  }

  const auth = readAuthFile();
  
  if (!auth[providerId]) {
    console.log(`Provider ${providerId} not found in auth file, nothing to remove`);
    return false;
  }

  delete auth[providerId];
  writeAuthFile(auth);
  console.log(`Removed provider auth: ${providerId}`);
  return true;
}

function getProviderAuth(providerId) {
  const auth = readAuthFile();
  return auth[providerId] || null;
}

function listProviderAuths() {
  const auth = readAuthFile();
  return Object.keys(auth);
}

export {
  readAuthFile,
  writeAuthFile,
  removeProviderAuth,
  getProviderAuth,
  listProviderAuths,
  getAuthFilePath as AUTH_FILE,
  getOpencodeDataDir as OPENCODE_DATA_DIR
};
