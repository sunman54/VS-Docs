import * as docx from 'docx-preview';

export async function renderDocx(
  buffer: ArrayBuffer,
  container: HTMLElement,
  isDarkMode: boolean = false
): Promise<void> {
  container.innerHTML = '';

  const docxContainer = document.createElement('div');
  docxContainer.id = 'docx-render-target';
  container.appendChild(docxContainer);

  await docx.renderAsync(buffer, docxContainer, undefined, {
    className: 'docx',
    inWrapper: true,
    ignoreWidth: false,
    ignoreHeight: false,
    ignoreFonts: false,
    breakPages: true,
    experimental: true,
    useBase64URL: true,
    trimXmlDeclaration: true,
  });

  // Apply smart color transformation for dark mode
  adjustDocxColorsForTheme(docxContainer, isDarkMode);
}

/**
 * Intelligent color adaptation:
 * In Word documents, dark text (black, #000000, windowtext) becomes invisible on dark backgrounds.
 * We detect near-black colors and adapt them, while strictly preserving intentional colorful
 * texts (red warning, blue hyperlinks, green notes, etc.).
 */
export function adjustDocxColorsForTheme(container: HTMLElement, isDarkMode: boolean) {
  const elements = container.querySelectorAll<HTMLElement>('*');

  elements.forEach((el) => {
    // 1. Text color adaptation
    if (isDarkMode) {
      const inlineColor = el.style.color;
      if (inlineColor) {
        if (!el.dataset.origColor) {
          el.dataset.origColor = inlineColor;
        }

        if (isNearBlackColor(inlineColor)) {
          el.style.color = '#e8eaed';
        }
      }

      // Check background color: If white/near-white background is set inline, clear or adapt
      const inlineBg = el.style.backgroundColor;
      if (inlineBg) {
        if (!el.dataset.origBg) {
          el.dataset.origBg = inlineBg;
        }
        if (isNearWhiteColor(inlineBg)) {
          el.style.backgroundColor = 'transparent';
        }
      }
    } else {
      // Revert to original colors if previously altered
      if (el.dataset.origColor !== undefined) {
        el.style.color = el.dataset.origColor;
      }
      if (el.dataset.origBg !== undefined) {
        el.style.backgroundColor = el.dataset.origBg;
      }
    }
  });
}

function isNearBlackColor(colorStr: string): boolean {
  const normalized = colorStr.trim().toLowerCase();
  if (['black', '#000', '#000000', 'windowtext', 'rgb(0, 0, 0)', 'rgba(0, 0, 0, 1)'].includes(normalized)) {
    return true;
  }
  // Parse rgb/rgba
  const rgbMatch = normalized.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (rgbMatch) {
    const r = parseInt(rgbMatch[1], 10);
    const g = parseInt(rgbMatch[2], 10);
    const b = parseInt(rgbMatch[3], 10);
    // Luminance check (standard formula)
    const luminance = 0.299 * r + 0.587 * g + 0.114 * b;
    return luminance < 60; // Very dark text
  }
  return false;
}

function isNearWhiteColor(colorStr: string): boolean {
  const normalized = colorStr.trim().toLowerCase();
  if (['white', '#fff', '#ffffff', 'rgb(255, 255, 255)', 'rgba(255, 255, 255, 1)'].includes(normalized)) {
    return true;
  }
  const rgbMatch = normalized.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (rgbMatch) {
    const r = parseInt(rgbMatch[1], 10);
    const g = parseInt(rgbMatch[2], 10);
    const b = parseInt(rgbMatch[3], 10);
    const luminance = 0.299 * r + 0.587 * g + 0.114 * b;
    return luminance > 220; // Very bright/white background
  }
  return false;
}


