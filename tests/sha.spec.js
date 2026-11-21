// The SHA hashes of src/docx/sha.js, which hash a password's 100,000 rounds for encryption.js
// without a promise per round: the same as Web Crypto's, for each hash and around the lengths
// where padding takes another block. Run in Node: nothing here needs a page.
import { webcrypto } from 'node:crypto';
import { test, expect } from '@playwright/test';
import { digest } from '../src/docx/sha.js';

const hex = (bytes) => Buffer.from(bytes).toString('hex');

for (const name of ['SHA-1', 'SHA-256', 'SHA-384', 'SHA-512']) {
  test(`${name} is Web Crypto’s, whatever the length`, async () => {
    for (const length of [0, 1, 3, 4, 55, 56, 63, 64, 65, 68, 111, 112, 119, 120, 127, 128, 129, 1000, 4100]) {
      const data = webcrypto.getRandomValues(new Uint8Array(length));
      expect(hex(digest(name, data)), `${name} of ${length} bytes`).toBe(hex(await webcrypto.subtle.digest(name, data)));
    }
  });
}

test('the known answers: "abc", as FIPS 180-4 gives them', () => {
  const abc = new TextEncoder().encode('abc');
  expect(hex(digest('SHA-1', abc))).toBe('a9993e364706816aba3e25717850c26c9cd0d89d');
  expect(hex(digest('SHA-256', abc))).toBe('ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad');
  expect(hex(digest('SHA-512', abc))).toBe('ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f');
});

test('100,000 rounds of SHA-512, as Word hashes a password, take well under a few seconds', () => {
  const buf = new Uint8Array(68);
  const start = performance.now();
  for (let i = 0; i < 100000; i++) buf.set(digest('SHA-512', buf), 4);
  expect(performance.now() - start).toBeLessThan(3000);
});
