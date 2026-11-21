// Password-encrypted Word documents (Word's File › Info › Protect Document › Encrypt with Password),
// decrypted and encrypted in the browser, on this device: the password and the document never leave
// it. Word 2010 and later encrypt with MS-OFFCRYPTO's Agile Encryption: the .docx package, encrypted
// with AES in 4096-byte segments under a random key, that key encrypted with one derived from the
// password (a salted hash, hashed again 100,000 times), and an HMAC over the encrypted package so a
// change to it is noticed; all of it in a compound file (see cfb.js). The browser's Web Crypto does
// the AES, the HMAC and single hashes; the password's rounds are hashed by sha.js, without a promise
// per round (which takes seconds in Firefox). Office 2007's Standard Encryption is opened too (and
// saved as Agile); the older RC4 kinds aren't.

import { readCfb, writeCfb } from './cfb.js';
import { digest } from './sha.js';

/** Thrown when a password is needed (`reason` 'needed') or the one given is wrong ('wrong'). */
export class PasswordError extends Error {
  constructor(reason, message) {
    super(message);
    this.name = 'PasswordError';
    this.reason = reason;
  }
}

const XMLNS_ENC = 'http://schemas.microsoft.com/office/2006/encryption';
const XMLNS_PASSWORD = 'http://schemas.microsoft.com/office/2006/keyEncryptor/password';
const SEGMENT = 4096;
const HASHES = { SHA1: 'SHA-1', 'SHA-1': 'SHA-1', SHA256: 'SHA-256', 'SHA-256': 'SHA-256', SHA384: 'SHA-384', 'SHA-384': 'SHA-384', SHA512: 'SHA-512', 'SHA-512': 'SHA-512' };
// The block keys MS-OFFCRYPTO hashes in for each use of the password's key and the document's key.
const BLOCK = {
  verifierInput: [0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79],
  verifierValue: [0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e],
  key: [0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6],
  hmacKey: [0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6],
  hmacValue: [0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33],
};

const subtle = () => {
  if (!globalThis.crypto?.subtle) throw new Error('This browser can’t decrypt or encrypt documents here (Web Crypto isn’t available on this page)');
  return globalThis.crypto.subtle;
};
const concat = (...parts) => {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of parts) {
    out.set(p, at);
    at += p.length;
  }
  return out;
};
const le32 = (n) => new Uint8Array(new Uint32Array([n]).buffer);
const hash = async (name, data) => new Uint8Array(await subtle().digest(name, data));
const b64 = (bytes) => {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
};
const unb64 = (s) => Uint8Array.from(atob((s || '').replace(/\s+/g, '')), (c) => c.charCodeAt(0));
const random = (n) => crypto.getRandomValues(new Uint8Array(n));
/** `bytes` made `n` long: cut, or padded with `fill` (0x36 for keys and IVs, as MS-OFFCRYPTO pads them). */
const sized = (bytes, n, fill = 0x36) => {
  if (bytes.length >= n) return bytes.slice(0, n);
  const out = new Uint8Array(n).fill(fill);
  out.set(bytes);
  return out;
};
const padTo = (bytes, block) => sized(bytes, Math.ceil(bytes.length / block) * block || block, 0);
const sameBytes = (a, b) => a.length === b.length && a.every((x, i) => x === b[i]);

// AES-CBC without padding: Web Crypto always pads (PKCS#7), so encrypting drops the padding block it
// adds, and decrypting adds the block that decrypts to a whole block of padding after the last one.
const aesKey = (raw, use) => subtle().importKey('raw', raw, 'AES-CBC', false, [use]);
async function aesEncrypt(raw, iv, data) {
  const out = new Uint8Array(await subtle().encrypt({ name: 'AES-CBC', iv }, await aesKey(raw, 'encrypt'), data));
  return out.subarray(0, data.length);
}
async function aesDecrypt(raw, iv, data) {
  if (!data.length || data.length % 16) throw new PasswordError('damaged', 'The encrypted document is damaged: its length isn’t whole blocks');
  const last = data.subarray(data.length - 16);
  const pad = new Uint8Array(await subtle().encrypt({ name: 'AES-CBC', iv: last }, await aesKey(raw, 'encrypt'), new Uint8Array(0)));
  const plain = await subtle().decrypt({ name: 'AES-CBC', iv }, await aesKey(raw, 'decrypt'), concat(data, pad));
  return new Uint8Array(plain);
}

/** The data spaces Word writes beside the encrypted package, saying it is encrypted (MS-OFFCRYPTO 2.1). */
function dataSpaces() {
  const u16 = (s) => {
    const out = new Uint8Array(s.length * 2);
    for (let i = 0; i < s.length; i++) out.set([s.charCodeAt(i) & 0xff, s.charCodeAt(i) >> 8], i * 2);
    return out;
  };
  // A length-prefixed UTF-16 string, padded to 4 bytes.
  const lp = (s) => {
    const body = u16(s);
    return concat(le32(body.length), body, new Uint8Array((4 - (body.length % 4)) % 4));
  };
  const v1 = concat(le32(1), le32(1), le32(1)); // reader, updater and writer versions 1.0
  const entry = concat(le32(1), le32(0), lp('EncryptedPackage'), lp('StrongEncryptionDataSpace'));
  // TransformLength counts itself, the type and the id, not the name after them.
  const id = concat(le32(1), lp('{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}'));
  const transform = concat(le32(id.length + 4), id, lp('Microsoft.Container.EncryptionTransform'));
  return new Map([
    ['\x06DataSpaces/Version', concat(lp('Microsoft.Container.DataSpaces'), v1)],
    ['\x06DataSpaces/DataSpaceMap', concat(le32(8), le32(1), le32(entry.length + 4), entry)],
    ['\x06DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace', concat(le32(8), le32(1), lp('StrongEncryptionTransform'))],
    ['\x06DataSpaces/TransformInfo/StrongEncryptionTransform/\x06Primary', concat(transform, v1, le32(0), le32(0), le32(0), le32(4))],
  ]);
}

/** The cipher settings of a keyData or encryptedKey element. */
function cipherOf(el) {
  const num = (a) => Number(el.getAttribute(a));
  const hashName = HASHES[(el.getAttribute('hashAlgorithm') || '').toUpperCase()];
  const ok = (el.getAttribute('cipherAlgorithm') || '').toUpperCase() === 'AES' && el.getAttribute('cipherChaining') === 'ChainingModeCBC' && hashName
    && [128, 192, 256].includes(num('keyBits')) && num('blockSize') === 16;
  if (!ok) {
    throw new PasswordError('unsupported', `This document is encrypted with ${el.getAttribute('cipherAlgorithm') || 'a cipher'} and ${el.getAttribute('hashAlgorithm') || 'a hash'}, which Intact doesn’t decrypt (it decrypts AES, as Word 2010 and later encrypt). Open it in Word.`);
  }
  return {
    hashName, keyBytes: num('keyBits') / 8, blockSize: num('blockSize'), hashSize: num('hashSize') || 0,
    salt: unb64(el.getAttribute('saltValue')),
  };
}

/** The password's key for each block key: the salted password hashed `spin` more times, then with the block key. */
async function passwordKeys(password, { hashName, salt, keyBytes }, spin, blocks) {
  const pw = new Uint8Array(password.length * 2);
  for (let i = 0; i < password.length; i++) pw.set([password.charCodeAt(i) & 0xff, password.charCodeAt(i) >> 8], i * 2);
  let h = await hash(hashName, concat(salt, pw));
  const buf = new Uint8Array(4 + h.length);
  const view = new DataView(buf.buffer);
  for (let i = 0; i < spin; i++) {
    view.setUint32(0, i, true);
    buf.set(h, 4);
    h = digest(hashName, buf); // synchronous: a promise per round would take seconds in Firefox
  }
  return Promise.all(blocks.map(async (b) => sized(await hash(hashName, concat(h, new Uint8Array(b))), keyBytes)));
}

/** The IV for segment or purpose `block` (a block key or a segment's number): the keyData salt hashed with it. */
const ivFor = async (c, block) => sized(await hash(c.hashName, concat(c.salt, block)), c.blockSize);

/** Whether `bytes` (a compound file) is a password-encrypted package, by its streams. */
export function isPasswordEncrypted(bytes) {
  try {
    const { streams } = readCfb(bytes);
    return streams.has('EncryptionInfo') && streams.has('EncryptedPackage');
  } catch {
    return false;
  }
}

/**
 * Decrypts a password-encrypted .docx (the compound file's `bytes`) with `password`: { zip (the
 * package's bytes), key (to encrypt the edited document again: see encrypt) }. Throws a
 * PasswordError: 'wrong' for a wrong password, 'unsupported' for an encryption Intact doesn't read,
 * 'damaged' when the package doesn't match its HMAC.
 */
export async function decrypt(bytes, password) {
  const { streams } = readCfb(bytes);
  const info = streams.get('EncryptionInfo');
  const pkg = streams.get('EncryptedPackage');
  if (!info || !pkg) throw new PasswordError('unsupported', 'This file doesn’t hold an encrypted document Intact can read');
  const major = info[0] | (info[1] << 8);
  const minor = info[2] | (info[3] << 8);
  if (major >= 2 && major <= 4 && minor === 2) return decryptStandard(info, pkg, password);
  if (!(major === 4 && minor === 4)) throw new PasswordError('unsupported', 'This document is encrypted in a way Intact doesn’t decrypt (an older kind than Word 2007’s). Open it in Word.');
  const xml = new DOMParser().parseFromString(new TextDecoder().decode(info.subarray(8)).replace(/^[^<]*/, ''), 'application/xml');
  const keyData = xml.getElementsByTagNameNS(XMLNS_ENC, 'keyData')[0];
  const encryptedKey = xml.getElementsByTagNameNS(XMLNS_PASSWORD, 'encryptedKey')[0];
  if (!keyData || !encryptedKey) throw new PasswordError('unsupported', 'This document is encrypted for a certificate, not a password, so Intact can’t open it. Open it in Word.');
  const data = cipherOf(keyData);
  const pw = cipherOf(encryptedKey);
  const spin = Number(encryptedKey.getAttribute('spinCount'));
  if (!(spin >= 0 && spin <= 10_000_000)) throw new PasswordError('unsupported', 'This document’s encryption settings can’t be read');
  const [kIn, kVal, kKey] = await passwordKeys(password, pw, spin, [BLOCK.verifierInput, BLOCK.verifierValue, BLOCK.key]);
  const iv = sized(pw.salt, pw.blockSize);
  const attr = (a) => unb64(encryptedKey.getAttribute(a));
  // The password is right when the verifier's hash decrypts to the hash of the decrypted verifier.
  const verifier = (await aesDecrypt(kIn, iv, attr('encryptedVerifierHashInput'))).subarray(0, pw.salt.length);
  const verifierHash = (await aesDecrypt(kVal, iv, attr('encryptedVerifierHashValue'))).subarray(0, pw.hashSize || undefined);
  const expected = await hash(pw.hashName, verifier);
  if (!sameBytes(expected, verifierHash.subarray(0, expected.length))) throw new PasswordError('wrong', 'That password isn’t right');
  const secret = (await aesDecrypt(kKey, iv, attr('encryptedKeyValue'))).subarray(0, data.keyBytes);

  // The package's HMAC, where it has one (Word always writes it): the encrypted package unchanged.
  const integrity = xml.getElementsByTagNameNS(XMLNS_ENC, 'dataIntegrity')[0];
  if (integrity) {
    const hmacKey = (await aesDecrypt(secret, await ivFor(data, new Uint8Array(BLOCK.hmacKey)), unb64(integrity.getAttribute('encryptedHmacKey')))).subarray(0, data.hashSize);
    const hmacValue = (await aesDecrypt(secret, await ivFor(data, new Uint8Array(BLOCK.hmacValue)), unb64(integrity.getAttribute('encryptedHmacValue')))).subarray(0, data.hashSize);
    const k = await subtle().importKey('raw', hmacKey, { name: 'HMAC', hash: data.hashName }, false, ['sign']);
    const actual = new Uint8Array(await subtle().sign('HMAC', k, pkg));
    if (!sameBytes(actual, hmacValue)) throw new PasswordError('damaged', 'The encrypted document has been changed or damaged since it was encrypted (its check doesn’t match), so it isn’t opened');
  }

  const size = Number(new DataView(pkg.buffer, pkg.byteOffset, 8).getBigUint64(0, true));
  if (size > pkg.length) throw new PasswordError('damaged', 'The encrypted document is damaged: it is shorter than it says');
  const out = new Uint8Array(Math.ceil(size / 16) * 16 + 16);
  let at = 0;
  for (let i = 0, from = 8; from < pkg.length && at < size; i++, from += SEGMENT) {
    const piece = pkg.subarray(from, Math.min(from + SEGMENT, pkg.length));
    const plain = await aesDecrypt(secret, await ivFor(data, le32(i)), piece.subarray(0, piece.length - (piece.length % 16)));
    out.set(plain.subarray(0, Math.min(plain.length, out.length - at)), at);
    at += plain.length;
  }
  const keyEncryptors = xml.getElementsByTagNameNS(XMLNS_ENC, 'keyEncryptors')[0];
  return {
    zip: out.slice(0, size),
    key: { secret, data: { ...data, salt: null }, keyEncryptors: new XMLSerializer().serializeToString(keyEncryptors) },
  };
}

/**
 * Standard Encryption (MS-OFFCRYPTO 2.3.4.5), as Office 2007 encrypts and LibreOffice still does:
 * AES in ECB mode over the whole package, its key from SHA-1 hashed 50,000 times. Such a document
 * is encrypted again as Word 2010 and later do (Agile), with the same password: `key` is null here,
 * and the caller makes one with keyForPassword.
 */
async function decryptStandard(info, pkg, password) {
  const v = new DataView(info.buffer, info.byteOffset, info.byteLength);
  const flags = v.getUint32(4, true);
  const headerSize = v.getUint32(8, true);
  const algId = v.getUint32(20, true);
  const hashId = v.getUint32(24, true);
  const keyBits = v.getUint32(28, true) || 128;
  // fAES set, fExternal not: AES with SHA-1 (RC4 "CryptoAPI" encryption isn't read).
  if (!(flags & 0x20) || (flags & 0x10) || ![0x660e, 0x660f, 0x6610].includes(algId) || ![0, 0x8004].includes(hashId)) {
    throw new PasswordError('unsupported', 'This document is encrypted in a way Intact doesn’t decrypt (RC4 or an external provider). Open it in Word.');
  }
  let at = 12 + headerSize;
  const saltSize = v.getUint32(at, true);
  const salt = info.slice(at + 4, at + 4 + saltSize);
  at += 4 + saltSize;
  const encVerifier = info.slice(at, at + 16);
  const hashSize = v.getUint32(at + 16, true);
  const encVerifierHash = info.slice(at + 20, at + 20 + 32);
  // The key: SHA-1 of salt and password, 50,000 more rounds, then with block 0, widened as CryptDeriveKey does.
  const pw = new Uint8Array(password.length * 2);
  for (let i = 0; i < password.length; i++) pw.set([password.charCodeAt(i) & 0xff, password.charCodeAt(i) >> 8], i * 2);
  let h = await hash('SHA-1', concat(salt, pw));
  const buf = new Uint8Array(24);
  const bv = new DataView(buf.buffer);
  for (let i = 0; i < 50000; i++) {
    bv.setUint32(0, i, true);
    buf.set(h, 4);
    h = digest('SHA-1', buf);
  }
  h = await hash('SHA-1', concat(h, le32(0)));
  const xor = (b) => new Uint8Array(64).map((_, i) => b ^ (h[i] || 0));
  const key = concat(await hash('SHA-1', xor(0x36)), await hash('SHA-1', xor(0x5c))).slice(0, keyBits / 8);
  const ecb = async (data) => {
    const plain = await aesDecrypt(key, new Uint8Array(16), data);
    for (let i = data.length - 16; i >= 16; i -= 16) for (let k = 0; k < 16; k++) plain[i + k] ^= data[i - 16 + k];
    return plain;
  };
  const verifier = await ecb(encVerifier);
  const verifierHash = (await ecb(encVerifierHash)).subarray(0, hashSize || 20);
  if (!sameBytes(await hash('SHA-1', verifier), verifierHash)) throw new PasswordError('wrong', 'That password isn’t right');
  const size = Number(new DataView(pkg.buffer, pkg.byteOffset, 8).getBigUint64(0, true));
  const body = pkg.subarray(8, 8 + Math.ceil(size / 16) * 16);
  if (size > pkg.length || body.length % 16) throw new PasswordError('damaged', 'The encrypted document is damaged: it is shorter than it says');
  return { zip: (await ecb(body)).slice(0, size), key: null, standard: true };
}

/**
 * A new key for encrypting a document with `password`, as Word 2010 and later make one: AES-256,
 * SHA-512, a 16-byte salt and 100,000 rounds. (The rounds are the slow part, done once.)
 */
export async function keyForPassword(password) {
  const spin = 100000;
  const pw = { hashName: 'SHA-512', keyBytes: 32, blockSize: 16, hashSize: 64, salt: random(16) };
  const secret = random(32);
  const [kIn, kVal, kKey] = await passwordKeys(password, pw, spin, [BLOCK.verifierInput, BLOCK.verifierValue, BLOCK.key]);
  const verifier = random(16);
  const enc = (k, data) => aesEncrypt(k, pw.salt, padTo(data, 16));
  const [hashInput, hashValue, keyValue] = await Promise.all([enc(kIn, verifier), hash('SHA-512', verifier).then((h) => enc(kVal, h)), enc(kKey, secret)]);
  const keyEncryptors = `<keyEncryptors xmlns="${XMLNS_ENC}"><keyEncryptor uri="${XMLNS_PASSWORD}">`
    + `<p:encryptedKey xmlns:p="${XMLNS_PASSWORD}" spinCount="${spin}" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512"`
    + ` saltValue="${b64(pw.salt)}" encryptedVerifierHashInput="${b64(hashInput)}" encryptedVerifierHashValue="${b64(hashValue)}" encryptedKeyValue="${b64(keyValue)}"/>`
    + '</keyEncryptor></keyEncryptors>';
  return { secret, data: { hashName: 'SHA-512', keyBytes: 32, blockSize: 16, hashSize: 64 }, keyEncryptors };
}

const HASH_NAMES = { 'SHA-1': 'SHA1', 'SHA-256': 'SHA256', 'SHA-384': 'SHA384', 'SHA-512': 'SHA512' };

/**
 * Encrypts a .docx package (`zip`, its bytes) with `key` (from decrypt or keyForPassword), as Word
 * does: a new salt for its segments each time, a new HMAC key, in a compound file with Word's data
 * spaces. Returns the compound file's bytes.
 */
export async function encrypt(zip, key) {
  const c = { ...key.data, salt: random(16) };
  const hashSize = c.hashSize || (await hash(c.hashName, new Uint8Array(0))).length;
  const segments = [];
  for (let i = 0, from = 0; from < zip.length; i++, from += SEGMENT) {
    segments.push(await aesEncrypt(key.secret, await ivFor(c, le32(i)), padTo(zip.subarray(from, from + SEGMENT), 16)));
  }
  const size = new Uint8Array(8);
  new DataView(size.buffer).setBigUint64(0, BigInt(zip.length), true);
  const pkg = concat(size, ...segments);
  const hmacKey = random(hashSize);
  const k = await subtle().importKey('raw', hmacKey, { name: 'HMAC', hash: c.hashName }, false, ['sign']);
  const hmacValue = new Uint8Array(await subtle().sign('HMAC', k, pkg));
  const encHmacKey = await aesEncrypt(key.secret, await ivFor(c, new Uint8Array(BLOCK.hmacKey)), padTo(hmacKey, c.blockSize));
  const encHmacValue = await aesEncrypt(key.secret, await ivFor(c, new Uint8Array(BLOCK.hmacValue)), padTo(hmacValue, c.blockSize));
  const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    + `<encryption xmlns="${XMLNS_ENC}" xmlns:p="${XMLNS_PASSWORD}" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">`
    + `<keyData saltSize="16" blockSize="${c.blockSize}" keyBits="${c.keyBytes * 8}" hashSize="${hashSize}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="${HASH_NAMES[c.hashName]}" saltValue="${b64(c.salt)}"/>`
    + `<dataIntegrity encryptedHmacKey="${b64(encHmacKey)}" encryptedHmacValue="${b64(encHmacValue)}"/>`
    + key.keyEncryptors.replace(/^<\?xml[^>]*>/, '').replace(/^<keyEncryptors xmlns="[^"]*"/, '<keyEncryptors')
    + '</encryption>';
  const info = concat(new Uint8Array([4, 0, 4, 0, 0x40, 0, 0, 0]), new TextEncoder().encode(xml));
  const streams = dataSpaces();
  streams.set('EncryptionInfo', info);
  streams.set('EncryptedPackage', pkg);
  return writeCfb(streams);
}
