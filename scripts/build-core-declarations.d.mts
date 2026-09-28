export interface DeclarationEntry {
  specifier: string;
  name: string;
  source: string;
}
export declare function publishedEntries(
  manifest: { exports: Record<string, unknown> },
  root?: string
): DeclarationEntry[];
export declare function packageName(specifier: string): string;
export declare function declarationCandidates(target: string): string[];
