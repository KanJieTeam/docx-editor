import type { TextMeasurer } from './semantic-records.ts';
import type { ResolvedRunStyle } from './run-style.ts';

export interface ResolvedPunctuationFace {
  readonly hash: string;
  readonly faceIndex: number;
  readonly family: string;
  readonly weight: number;
  readonly style: string;
}

type Provider = (style: ResolvedRunStyle, text: string) => ResolvedPunctuationFace | undefined;
const resolvedFace = Symbol('resolved-source-punctuation-face');
type WithFace = TextMeasurer & { readonly [resolvedFace]?: Provider };

// A private enumerable symbol survives host measurer spreads. No public style
// field or file-supplied font name can create this measurement capability.
export function withResolvedPunctuationFace(
  measurer: TextMeasurer,
  provider: Provider
): TextMeasurer {
  return { ...measurer, [resolvedFace]: provider } as WithFace;
}

export function resolvedPunctuationFace(
  measurer: TextMeasurer,
  style: ResolvedRunStyle,
  text: string
): ResolvedPunctuationFace | undefined {
  return (measurer as WithFace)[resolvedFace]?.(style, text);
}
