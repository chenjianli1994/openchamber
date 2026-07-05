import os from 'os';
import path from 'path';

const DEFAULT_OPENCODE_DATA_DIR = path.join(os.homedir(), '.local', 'share', 'opencode');

function pickFirstNonEmpty(values) {
  for (const value of values) {
    if (typeof value !== 'string') {
      continue;
    }

    const trimmed = value.trim();
    if (trimmed) {
      return trimmed;
    }
  }

  return '';
}

export function getLangCoderOpencodeDataDir(env = process.env) {
  const configured = pickFirstNonEmpty([
    env.LANGCODER_OPENCODE_DATA_DIR,
    env.OPENCODE_DATA_DIR,
  ]);

  return configured || DEFAULT_OPENCODE_DATA_DIR;
}

export function getLangCoderOpencodeAuthFilePath(env = process.env) {
  return path.join(getLangCoderOpencodeDataDir(env), 'auth.json');
}
