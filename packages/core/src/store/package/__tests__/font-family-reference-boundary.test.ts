import { describe, expect, test } from 'bun:test';
import { fontFamilyAlternatives, validFontFamilyReference } from '../font-family-reference.js';

describe('font family reference Unicode bounds', () => {
  test('keeps the existing 64-code-point limit for single names', () => {
    const name = '\u{20000}'.repeat(64);
    expect(name.length).toBe(128);
    expect(validFontFamilyReference(name)).toBe(name);
    expect(fontFamilyAlternatives(name)).toEqual([name]);
    expect(validFontFamilyReference(name + '\u{20000}')).toBeNull();
    expect(fontFamilyAlternatives(name + '\u{20000}')).toEqual([]);
  });

  test('counts separators and every candidate within the whole-reference bound', () => {
    const reference = '\u{20000}'.repeat(62) + ';B';
    expect([...reference]).toHaveLength(64);
    expect(validFontFamilyReference(reference)).toBe(reference);
    expect(fontFamilyAlternatives(reference)).toEqual(['\u{20000}'.repeat(62), 'B']);
    expect(validFontFamilyReference(reference + 'C')).toBeNull();
    expect(fontFamilyAlternatives(reference + 'C')).toEqual([]);
  });

  test('keeps the pre-parse memory bound and name safety checks', () => {
    expect(fontFamilyAlternatives('\u{20000}'.repeat(10000))).toEqual([]);
    expect(fontFamilyAlternatives('\u{20000};bad\nname')).toEqual([]);
    expect(fontFamilyAlternatives('\u{20000};bad"name')).toEqual([]);
    expect(fontFamilyAlternatives('\u{20000};bad/name')).toEqual([]);
  });
});
