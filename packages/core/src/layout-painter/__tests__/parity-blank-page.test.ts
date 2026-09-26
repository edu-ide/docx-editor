/**
 * The blank page an odd/even-page section break inserts is painted empty:
 * no header, footer, watermark or page border (upstream 6794f4d3, #978).
 * Any other page keeps its furniture.
 */
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { PAGE_CLASS_NAMES, renderPage, type HeaderFooterContent } from '../renderPage';
import { WATERMARK_LAYER_CLASS } from '../renderWatermark';
import type { Page, ParagraphBlock, ParagraphMeasure } from '../../layout-engine/types';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

function makePage(parityBlank?: boolean): Page {
  return {
    number: 2,
    fragments: [],
    margins: { top: 96, right: 96, bottom: 96, left: 96, header: 48, footer: 48 },
    size: { w: 816, h: 1056 },
    ...(parityBlank ? { parityBlank } : {}),
  };
}

function storyContent(text: string): HeaderFooterContent {
  const block: ParagraphBlock = {
    kind: 'paragraph',
    id: `story-${text}`,
    runs: [{ kind: 'text', text }],
    attrs: {},
  };
  const measure: ParagraphMeasure = {
    kind: 'paragraph',
    lines: [
      {
        fromRun: 0,
        fromChar: 0,
        toRun: 0,
        toChar: text.length,
        width: 40,
        ascent: 12,
        descent: 4,
        lineHeight: 16,
      },
    ],
    totalHeight: 16,
  };
  return { blocks: [block], measures: [measure], height: 16, visualTop: 0, visualBottom: 16 };
}

function paint(page: Page): HTMLElement {
  return renderPage(
    page,
    { pageNumber: page.number, totalPages: 3, section: 'body' },
    {
      document,
      headerContent: storyContent('HEADER'),
      footerContent: storyContent('FOOTER'),
      watermark: {
        kind: 'text',
        text: 'DRAFT',
        font: 'Calibri',
        color: '#C0C0C0',
        semitransparent: true,
        layout: 'diagonal',
      },
      pageBorders: {
        display: 'allPages',
        offsetFrom: 'page',
        top: { style: 'single', size: 8, space: 24, color: { rgb: '000000' } },
      },
    }
  );
}

describe('blank parity page', () => {
  test('paints no header, footer, watermark or page border', () => {
    const el = paint(makePage(true));
    expect(el.dataset.parityBlank).toBe('true');
    expect(el.querySelector(`.${PAGE_CLASS_NAMES.header}`)).toBeNull();
    expect(el.querySelector(`.${PAGE_CLASS_NAMES.footer}`)).toBeNull();
    expect(el.querySelector(`.${WATERMARK_LAYER_CLASS}`)).toBeNull();
    expect(el.querySelector('.layout-page-border')).toBeNull();
    expect(el.querySelector(`.${PAGE_CLASS_NAMES.content}`)).not.toBeNull();
    expect(el.textContent).toBe('');
    expect(el.dataset.pageNumber).toBe('2');
  });

  test('an ordinary empty page keeps its furniture', () => {
    const el = paint(makePage());
    expect(el.dataset.parityBlank).toBeUndefined();
    expect(el.querySelector(`.${PAGE_CLASS_NAMES.header}`)?.textContent).toContain('HEADER');
    expect(el.querySelector(`.${PAGE_CLASS_NAMES.footer}`)?.textContent).toContain('FOOTER');
    expect(el.querySelector(`.${WATERMARK_LAYER_CLASS}`)).not.toBeNull();
    expect(el.querySelector('.layout-page-border')).not.toBeNull();
  });
});
