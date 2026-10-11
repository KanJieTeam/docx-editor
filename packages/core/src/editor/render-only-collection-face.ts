import { HARD_MAX_FONT_BYTES } from '../layout/font-resource.ts';
import { materializeCollectionFace } from './font-collection-face.ts';

/** A trusted host decision. Document font data must not set this policy. */
export interface RenderOnlyCollectionPolicy {
  readonly allowUnsignedDerivedFaces: boolean;
  readonly requireSignedFonts: boolean;
}

export interface RenderOnlyCollectionFace {
  readonly bytes: Uint8Array;
  readonly faceIndex: number;
  /** This derived file does not carry the source collection signature. */
  readonly signature: 'unsigned';
  readonly sourceCollectionSignaturePresent: boolean;
  /** Structural checks do not authenticate a digital signature. */
  readonly sourceSignatureValidated: false;
}

function requireValue(value: unknown, message: string): asserts value {
  if (!value) throw new Error(message);
}

/**
 * Derive an unsigned rendering file only after an explicit host decision.
 * Keep the admitted source collection intact. Never represent this output as signed.
 * Individual signed faces remain unsupported by the selected-face materializer.
 */
export function materializeRenderOnlyCollectionFace(
  bytes: Uint8Array,
  faceIndex: number,
  policy: RenderOnlyCollectionPolicy
): RenderOnlyCollectionFace {
  requireValue(
    policy?.allowUnsignedDerivedFaces === true && policy.requireSignedFonts === false,
    'Unsigned rendering projection is not permitted by the host'
  );
  requireValue(
    bytes instanceof Uint8Array && bytes.length >= 12 && bytes.length <= HARD_MAX_FONT_BYTES,
    'Unsupported font byte length'
  );
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const range = (offset: number, length: number) => {
    requireValue(
      Number.isSafeInteger(offset) &&
        Number.isSafeInteger(length) &&
        offset >= 0 &&
        length >= 0 &&
        offset <= bytes.length - length,
      'Font range exceeds source'
    );
  };
  requireValue(view.getUint32(0) === 0x74746366, 'Rendering projection requires a collection');
  const version = view.getUint32(4),
    count = view.getUint32(8);
  requireValue(version === 0x00010000 || version === 0x00020000, 'Unsupported TTC version');
  requireValue(
    Number.isSafeInteger(faceIndex) && faceIndex >= 0 && faceIndex < count,
    'Collection face absent'
  );
  const headerEnd = 12 + count * 4 + (version === 0x00020000 ? 12 : 0);
  range(0, headerEnd);
  const signatureFields = 12 + count * 4;
  let hasSignature = false;
  if (version === 0x00020000) {
    const tag = view.getUint32(signatureFields),
      length = view.getUint32(signatureFields + 4),
      offset = view.getUint32(signatureFields + 8);
    hasSignature = tag !== 0 || length !== 0 || offset !== 0;
    if (hasSignature) {
      requireValue(
        tag === 0x44534947 && length >= 8 && offset >= headerEnd && offset % 4 === 0,
        'Invalid collection signature directory'
      );
      range(offset, length);
      requireValue(
        bytes.length - offset - length <= 3,
        'Collection signature is not the final table'
      );
      for (let p = offset + length; p < bytes.length; p++)
        requireValue(bytes[p] === 0, 'Invalid signature padding');
      requireValue(
        view.getUint32(offset) === 1 && (view.getUint16(offset + 6) & ~1) === 0,
        'Invalid collection signature header'
      );
      const signatures = view.getUint16(offset + 4),
        directoryLength = 8 + signatures * 12;
      requireValue(directoryLength <= length, 'Truncated signature records');
      for (let i = 0; i < signatures; i++) {
        const entry = offset + 8 + i * 12;
        const format = view.getUint32(entry),
          size = view.getUint32(entry + 4),
          start = view.getUint32(entry + 8);
        requireValue(
          format === 1 && size >= 8 && start >= directoryLength && start <= length - size,
          'Invalid collection signature block'
        );
        requireValue(
          view.getUint16(offset + start) === 0 &&
            view.getUint16(offset + start + 2) === 0 &&
            view.getUint32(offset + start + 4) <= size - 8,
          'Invalid signature payload bounds'
        );
      }
      const base = view.getUint32(12 + faceIndex * 4);
      range(base, 12);
      const tables = view.getUint16(base + 4),
        directoryEnd = base + 12 + tables * 16;
      range(base, 12 + tables * 16);
      requireValue(
        directoryEnd <= offset || base >= offset + length,
        'Signature overlaps selected face directory'
      );
      for (let i = 0; i < tables; i++) {
        const entry = base + 12 + i * 16,
          tableOffset = view.getUint32(entry + 8),
          tableLength = view.getUint32(entry + 12);
        range(tableOffset, tableLength);
        requireValue(
          tableLength === 0 ||
            tableOffset + tableLength <= offset ||
            tableOffset >= offset + length,
          'Signature overlaps selected face table'
        );
      }
    }
  }
  // The old materializer keeps its refusal contract. This private temporary copy
  // describes a different, unsigned rendering input. The source is never changed.
  const renderingInput = hasSignature ? bytes.slice() : bytes;
  if (hasSignature)
    new Uint8Array(renderingInput.buffer, renderingInput.byteOffset + signatureFields, 12).fill(0);
  return Object.freeze({
    bytes: materializeCollectionFace(renderingInput, faceIndex),
    faceIndex,
    signature: 'unsigned',
    sourceCollectionSignaturePresent: hasSignature,
    sourceSignatureValidated: false,
  });
}
