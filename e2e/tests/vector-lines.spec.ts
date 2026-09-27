import { test, expect, type Page } from '@playwright/test';
import { EditorPage } from '../helpers/editor-page';

/**
 * Upstream #972 — straight lines and a horizontal rule drawn as `wps:wsp`
 * shapes must paint their stroke. Before, each painted as an empty picture
 * box, and a zero-width or zero-height line took a phantom 100px on that axis
 * (pushing its line of text down by 100px).
 *
 * The fixture gives every drawing a `wp:docPr descr`, so its painted image is
 * found by alt text.
 */
const FIXTURE = 'fixtures/vector-lines-and-rules.docx';

/** Layout size in px (at zoom 1) that each drawing must paint at. */
const EXPECTED = {
  'Vertical line': { width: 2, height: 96 },
  'Horizontal line': { width: 192, height: 4 / 3 },
  'Diagonal line': { width: 96, height: 48 },
  'Full-width rule': { width: 624, height: 2 },
  'Margin connector': { width: 4 / 3, height: 144 },
} as const;
type DrawingName = keyof typeof EXPECTED;

interface PaintedDrawing {
  found: boolean;
  src: string;
  width: number;
  height: number;
  /** The browser decoded the source (an SVG) into an image with a size. */
  decoded: boolean;
  /** Pixels the decoded image actually inks when drawn to a canvas. */
  inkPixels: number;
}

async function loadFixture(page: Page): Promise<void> {
  const editor = new EditorPage(page);
  await editor.goto();
  await editor.waitForReady();
  await editor.loadDocxFile(FIXTURE);
  // The demo document paints first; wait for this fixture's own drawings.
  await page.waitForSelector('[data-page-number="1"] img[alt="Full-width rule"]', {
    state: 'attached',
  });
  await page.waitForSelector('[data-page-number="1"] img[alt="Margin connector"]', {
    state: 'attached',
  });
}

/** Page zoom: painted px per layout px (the page is 816px wide at zoom 1). */
async function zoomOf(page: Page): Promise<number> {
  const box = await page.locator('[data-page-number="1"]').boundingBox();
  if (!box) throw new Error('page 1 is not painted');
  return box.width / 816;
}

async function paint(page: Page, name: DrawingName): Promise<PaintedDrawing> {
  return page.evaluate(async (alt) => {
    const img = document.querySelector<HTMLImageElement>(
      `[data-page-number="1"] img[alt="${alt}"]`
    );
    if (!img) return { found: false, src: '', width: 0, height: 0, decoded: false, inkPixels: 0 };
    let decoded = true;
    try {
      await img.decode();
    } catch {
      decoded = false;
    }
    const rect = img.getBoundingClientRect();
    let inkPixels = 0;
    if (decoded && img.naturalWidth > 0 && img.naturalHeight > 0) {
      const canvas = document.createElement('canvas');
      canvas.width = Math.max(1, Math.ceil(rect.width));
      canvas.height = Math.max(1, Math.ceil(rect.height));
      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        const { data } = ctx.getImageData(0, 0, canvas.width, canvas.height);
        for (let i = 3; i < data.length; i += 4) {
          if (data[i]! > 0) inkPixels++;
        }
      }
    }
    return {
      found: true,
      src: img.getAttribute('src') ?? '',
      width: rect.width,
      height: rect.height,
      decoded: decoded && img.naturalWidth > 0 && img.naturalHeight > 0,
      inkPixels,
    };
  }, name);
}

test.describe('Lines and rules drawn as shapes (#972)', () => {
  test.beforeEach(async ({ page }) => {
    await loadFixture(page);
  });

  for (const name of Object.keys(EXPECTED) as DrawingName[]) {
    test(`${name} paints its stroke at its own size`, async ({ page }) => {
      const zoom = await zoomOf(page);
      const painted = await paint(page, name);

      expect(painted.found).toBe(true);
      expect(painted.src).not.toBe('');
      expect(painted.decoded).toBe(true);
      expect(painted.inkPixels).toBeGreaterThan(0);

      // A non-zero stroke extent on both axes, and no phantom 100px axis.
      const width = painted.width / zoom;
      const height = painted.height / zoom;
      expect(width).toBeGreaterThan(0);
      expect(height).toBeGreaterThan(0);
      expect(Math.abs(width - 100)).toBeGreaterThan(1);
      expect(Math.abs(height - 100)).toBeGreaterThan(1);
      expect(width).toBeCloseTo(EXPECTED[name].width, 0);
      expect(height).toBeCloseTo(EXPECTED[name].height, 0);
    });
  }

  test('no empty picture box is painted', async ({ page }) => {
    const empty = await page.evaluate(() =>
      [...document.querySelectorAll<HTMLImageElement>('[data-page-number="1"] img')]
        .filter((img) => !img.getAttribute('src'))
        .map((img) => img.alt)
    );
    expect(empty).toEqual([]);
  });

  test('a horizontal line keeps its line at text height', async ({ page }) => {
    const heights = await page.evaluate(() => {
      const lineHeightOf = (marker: string) => {
        const paragraph = [
          ...document.querySelectorAll<HTMLElement>('[data-page-number="1"] .layout-paragraph'),
        ].find((el) => el.textContent?.includes(marker));
        const line = paragraph?.querySelector<HTMLElement>('.layout-line');
        return line ? line.getBoundingClientRect().height : -1;
      };
      return {
        horizontal: lineHeightOf('sits on this line'),
        vertical: lineHeightOf('stands between these words'),
        plain: lineHeightOf('Plain paragraph before'),
      };
    });
    const zoom = await zoomOf(page);

    expect(heights.plain).toBeGreaterThan(0);
    // The zero-height line is as tall as its stroke, well inside the text line.
    expect(Math.abs(heights.horizontal - heights.plain)).toBeLessThan(1);
    // The vertical line grows its line to its own 96px length, not past it.
    expect(heights.vertical / zoom).toBeGreaterThanOrEqual(96);
    expect(heights.vertical / zoom).toBeLessThan(96 + heights.plain / zoom);
  });
});
