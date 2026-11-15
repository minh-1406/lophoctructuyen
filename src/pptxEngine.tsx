import React, { useEffect, useRef, useState } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { RotateCcw, SkipBack, SkipForward, Layers, Sparkles, Volume2 } from 'lucide-react';

export interface PptxTransition {
  type: 'fade' | 'push' | 'wipe' | 'zoom' | 'split' | 'cover' | 'wheel' | 'flip' | 'cut' | 'none';
  dir?: 'l' | 'r' | 'u' | 'd';
  durationMs: number;
  label?: string;
}

export interface PptxAnimStep {
  spid: string;
  presetClass: 'entr' | 'emph' | 'exit' | 'path';
  effect: 'appear' | 'fade' | 'fly' | 'zoom' | 'wipe' | 'float' | 'bounce' | 'pulse' | 'spin' | 'fadeOut' | 'motionPath';
  dir?: 'l' | 'r' | 'u' | 'd';
  motionDxPct?: number;
  motionDyPct?: number;
  trigger: 'onClick' | 'withPrev' | 'afterPrev';
  durationMs: number;
  delayMs: number;
  clickGroup: number;
}

export interface PptxTextRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSizePt?: number;
  color?: string;
}

export interface PptxParagraph {
  align?: 'left' | 'center' | 'right' | 'justify';
  bullet?: boolean;
  runs: PptxTextRun[];
}

export interface PptxTableCell {
  text: string;
  paragraphs: PptxParagraph[];
  fill?: string;
  colSpan?: number;
  rowSpan?: number;
}

export interface PptxElement {
  id: string;
  kind: 'shape' | 'pic' | 'table';
  xPct: number;
  yPct: number;
  wPct: number;
  hPct: number;
  rotDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
  geom?: string;
  fill?: string;
  stroke?: string;
  strokeWidthPx?: number;
  paragraphs?: PptxParagraph[];
  valign?: 'top' | 'middle' | 'bottom';
  fontScale?: number;
  imgId?: string;
  vidId?: string;
  audioId?: string;
  tableRows?: PptxTableCell[][];
}

export interface PptxSlideData {
  widthPx: number;
  heightPx: number;
  background?: string;
  bgImgId?: string;
  elements: PptxElement[];
  transition?: PptxTransition;
  animations: PptxAnimStep[];
  audioIds?: string[];
}

// XML localName helpers (namespace-prefix independent)
const childEls = (parent: Element | Document | null | undefined, local?: string): Element[] => {
  if (!parent) return [];
  const res: Element[] = [];
  for (let i = 0; i < parent.children.length; i++) {
    const c = parent.children[i];
    if (!local || c.localName === local) res.push(c);
  }
  return res;
};

const firstChild = (parent: Element | null | undefined, local: string): Element | null => {
  if (!parent) return null;
  for (let i = 0; i < parent.children.length; i++) {
    if (parent.children[i].localName === local) return parent.children[i];
  }
  return null;
};

const descEls = (parent: Element | Document | null | undefined, local: string): Element[] => {
  if (!parent) return [];
  const all = parent.getElementsByTagName('*');
  const res: Element[] = [];
  for (let i = 0; i < all.length; i++) {
    if (all[i].localName === local) res.push(all[i]);
  }
  return res;
};

const getRelAttr = (el: Element | null | undefined, localAttr: string): string => {
  if (!el) return '';
  return (
    el.getAttribute(`r:${localAttr}`) ||
    el.getAttribute(localAttr) ||
    el.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', localAttr) ||
    ''
  );
};

const parseXmlDoc = (xml: string) => new DOMParser().parseFromString(xml, 'text/xml');

const resolveZipPath = (baseDir: string, target: string): string => {
  if (!target) return '';
  if (/^https?:\/\//i.test(target) || target.startsWith('data:')) return target;
  if (target.startsWith('/')) return target.slice(1);
  const parts = (baseDir ? baseDir.split('/') : []).filter(Boolean);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.' && seg !== '') parts.push(seg);
  }
  return parts.join('/');
};

const parseRelsMap = async (zip: any, relsPath: string, baseDir: string): Promise<Record<string, { target: string; type: string }>> => {
  const f = zip.file(relsPath);
  if (!f) return {};
  const doc = parseXmlDoc(await f.async('string'));
  const map: Record<string, { target: string; type: string }> = {};
  for (const rel of descEls(doc, 'Relationship')) {
    const id = rel.getAttribute('Id') || '';
    const target = rel.getAttribute('Target') || '';
    const type = rel.getAttribute('Type') || '';
    if (id && target) {
      map[id] = { target: resolveZipPath(baseDir, target), type };
    }
  }
  return map;
};

// Color utilities
const hexToRgb = (hex: string): [number, number, number] => {
  const clean = hex.replace('#', '');
  const n = parseInt(clean.length === 3 ? clean.split('').map(c => c + c).join('') : clean, 16);
  if (Number.isNaN(n)) return [0, 0, 0];
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
};

const rgbToHsl = (r: number, g: number, b: number): [number, number, number] => {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h /= 6;
  }
  return [h, s, l];
};

const hslToRgb = (h: number, s: number, l: number): [number, number, number] => {
  if (s === 0) {
    const v = Math.round(l * 255);
    return [v, v, v];
  }
  const hue2rgb = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [
    Math.round(hue2rgb(p, q, h + 1 / 3) * 255),
    Math.round(hue2rgb(p, q, h) * 255),
    Math.round(hue2rgb(p, q, h - 1 / 3) * 255),
  ];
};

const applyColorMods = (hex: string, colorEl: Element): string => {
  let [r, g, b] = hexToRgb(hex);
  let alpha = 1;

  for (const mod of childEls(colorEl)) {
    const val = parseInt(mod.getAttribute('val') || '0', 10) / 100000;
    if (mod.localName === 'alpha') {
      alpha = Math.max(0, Math.min(1, val));
    } else if (mod.localName === 'tint') {
      r = Math.round(r + (255 - r) * (1 - val));
      g = Math.round(g + (255 - g) * (1 - val));
      b = Math.round(b + (255 - b) * (1 - val));
    } else if (mod.localName === 'shade') {
      r = Math.round(r * val);
      g = Math.round(g * val);
      b = Math.round(b * val);
    } else if (mod.localName === 'lumMod' || mod.localName === 'lumOff') {
      const [h, s, l] = rgbToHsl(r, g, b);
      const nextL = mod.localName === 'lumMod' ? Math.min(1, l * val) : Math.min(1, Math.max(0, l + val));
      [r, g, b] = hslToRgb(h, s, nextL);
    }
  }

  if (alpha < 0.99) {
    return `rgba(${r}, ${g}, ${b}, ${alpha.toFixed(2)})`;
  }
  return `#${[r, g, b].map(v => v.toString(16).padStart(2, '0')).join('')}`;
};

const isDarkColor = (cssColor?: string): boolean => {
  if (!cssColor) return false;
  if (cssColor.startsWith('#')) {
    const [r, g, b] = hexToRgb(cssColor);
    const lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    return lum < 0.45;
  }
  const m = cssColor.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (m) {
    const lum = (0.299 * Number(m[1]) + 0.587 * Number(m[2]) + 0.114 * Number(m[3])) / 255;
    return lum < 0.45;
  }
  return false;
};

const resolveColorEl = (container: Element | null | undefined, themeColors: Record<string, string>): string | undefined => {
  if (!container) return undefined;
  const srgb = firstChild(container, 'srgbClr');
  if (srgb) {
    const val = srgb.getAttribute('val');
    if (val) return applyColorMods('#' + val, srgb);
  }
  const scheme = firstChild(container, 'schemeClr');
  if (scheme) {
    const val = scheme.getAttribute('val') || '';
    const baseHex = themeColors[val] || themeColors.dk1 || '#1e293b';
    return applyColorMods(baseHex, scheme);
  }
  const sys = firstChild(container, 'sysClr');
  if (sys) {
    const val = sys.getAttribute('lastClr') || sys.getAttribute('val');
    if (val && /^[0-9a-fA-F]{6}$/.test(val)) return applyColorMods('#' + val, sys);
  }
  const prst = firstChild(container, 'prstClr');
  if (prst) {
    const val = prst.getAttribute('val');
    if (val === 'white') return '#ffffff';
    if (val === 'black') return '#000000';
    return val || undefined;
  }
  return undefined;
};

const extractFillStyle = (
  container: Element | null | undefined,
  themeColors: Record<string, string>
): { fill?: string; noFill?: boolean } => {
  if (!container) return {};
  if (firstChild(container, 'noFill')) return { noFill: true };
  const solid = firstChild(container, 'solidFill');
  if (solid) {
    const c = resolveColorEl(solid, themeColors);
    if (c) return { fill: c };
  }
  const grad = firstChild(container, 'gradFill');
  if (grad) {
    const gsLst = firstChild(grad, 'gsLst');
    const stops = childEls(gsLst, 'gs')
      .map(gs => {
        const pos = Math.round((parseInt(gs.getAttribute('pos') || '0', 10) / 100000) * 100);
        const c = resolveColorEl(gs, themeColors) || '#ffffff';
        return { pos, c };
      })
      .sort((a, b) => a.pos - b.pos);
    if (stops.length >= 2) {
      const lin = firstChild(grad, 'lin');
      const ang = lin ? Math.round(parseInt(lin.getAttribute('ang') || '5400000', 10) / 60000) + 90 : 180;
      return { fill: `linear-gradient(${ang}deg, ${stops.map(s => `${s.c} ${s.pos}%`).join(', ')})` };
    }
    if (stops.length === 1) return { fill: stops[0].c };
  }
  return {};
};

const parseThemeColors = async (zip: any): Promise<Record<string, string>> => {
  const defaults: Record<string, string> = {
    dk1: '#0f172a', lt1: '#ffffff', dk2: '#334155', lt2: '#f1f5f9',
    accent1: '#7c3aed', accent2: '#ec4899', accent3: '#10b981',
    accent4: '#f59e0b', accent5: '#3b82f6', accent6: '#6366f1',
    tx1: '#0f172a', bg1: '#ffffff', tx2: '#334155', bg2: '#f1f5f9',
    phClr: '#7c3aed',
  };
  const themeKey = Object.keys(zip.files).find(k => /^ppt\/theme\/theme\d+\.xml$/.test(k));
  if (!themeKey) return defaults;
  try {
    const doc = parseXmlDoc(await zip.file(themeKey).async('string'));
    const scheme = descEls(doc, 'clrScheme')[0];
    if (!scheme) return defaults;
    for (const c of childEls(scheme)) {
      const srgb = firstChild(c, 'srgbClr');
      const sys = firstChild(c, 'sysClr');
      const val = srgb?.getAttribute('val') || sys?.getAttribute('lastClr');
      if (val) defaults[c.localName] = '#' + val;
    }
    defaults.tx1 = defaults.dk1;
    defaults.bg1 = defaults.lt1;
    defaults.tx2 = defaults.dk2;
    defaults.bg2 = defaults.lt2;
  } catch {}
  return defaults;
};

const parseTransition = (doc: Document): PptxTransition | undefined => {
  const transEls = descEls(doc, 'transition');
  if (!transEls.length) return undefined;

  for (const transEl of transEls) {
    const spd = transEl.getAttribute('spd') || 'med';
    const durAttr = transEl.getAttribute('dur') || transEl.getAttribute('p14:dur');
    const durationMs = durAttr
      ? Math.max(150, Math.min(2500, parseInt(durAttr, 10) || 600))
      : spd === 'fast'
      ? 350
      : spd === 'slow'
      ? 1000
      : 600;

    const inspectChildren = (el: Element): Element[] => {
      const out: Element[] = [];
      for (const c of childEls(el)) {
        if (c.localName === 'AlternateContent') {
          const choice = firstChild(c, 'Choice');
          const fallback = firstChild(c, 'Fallback');
          if (choice) out.push(...childEls(choice));
          if (fallback) out.push(...childEls(fallback));
        } else if (c.localName !== 'sndAc' && c.localName !== 'extLst') {
          out.push(c);
        } else if (c.localName === 'extLst') {
          for (const ext of childEls(c, 'ext')) {
            out.push(...childEls(ext));
          }
        }
      }
      return out;
    };

    const candidates = inspectChildren(transEl);
    for (const el of candidates) {
      const tag = el.localName;
      const rawDir = el.getAttribute('dir') || 'l';
      const dir: 'l' | 'r' | 'u' | 'd' =
        rawDir === 'r' || rawDir === 'u' || rawDir === 'd' ? rawDir : 'l';

      if (['fade', 'dissolve', 'flash', 'morph', 'prism', 'glitter', 'honeycomb', 'ripple', 'vortex', 'shred'].includes(tag)) {
        return { type: 'fade', durationMs, label: tag === 'morph' ? 'Biến đổi (Morph)' : 'Mờ dần (Fade)' };
      }
      if (['push', 'pan', 'conveyor', 'comb', 'wind'].includes(tag)) {
        return { type: 'push', dir, durationMs, label: 'Đẩy (Push)' };
      }
      if (['wipe', 'strips', 'blinds', 'bars', 'checker', 'drape'].includes(tag)) {
        return { type: 'wipe', dir, durationMs, label: 'Quét (Wipe)' };
      }
      if (['zoom', 'box', 'flythrough', 'reveal', 'warp'].includes(tag)) {
        return { type: 'zoom', durationMs, label: 'Thu phóng (Zoom)' };
      }
      if (['split', 'doors', 'curtains'].includes(tag)) {
        return { type: 'split', durationMs, label: 'Chia đôi (Split)' };
      }
      if (['cover', 'pull', 'uncover', 'wedge', 'gallery'].includes(tag)) {
        return { type: 'cover', dir, durationMs, label: 'Trượt phủ (Cover)' };
      }
      if (['wheel', 'circle', 'diamond', 'plus', 'newsflash'].includes(tag)) {
        return { type: 'wheel', durationMs, label: 'Xoay tròn (Wheel)' };
      }
      if (['flip', 'cube', 'rotate', 'orbit', 'fallOver', 'switch', 'pageCurlDouble', 'pageCurlSingle'].includes(tag)) {
        return { type: 'flip', dir, durationMs, label: 'Lật 3D (Flip)' };
      }
      if (tag === 'cut') {
        return { type: 'cut', durationMs: 120, label: 'Cắt nhanh (Cut)' };
      }
    }
  }
  return undefined;
};

const parseTimingAnimations = (doc: Document): PptxAnimStep[] => {
  const timing = descEls(doc, 'timing')[0];
  if (!timing) return [];

  const steps: PptxAnimStep[] = [];
  let clickGroup = 0;

  const mapPreset = (
    presetClass: 'entr' | 'emph' | 'exit' | 'path',
    presetID: number,
    presetSubtype: number,
    filterStr: string
  ): { effect: PptxAnimStep['effect']; dir?: 'l' | 'r' | 'u' | 'd' } => {
    if (presetClass === 'path') return { effect: 'motionPath' };
    if (presetClass === 'exit') return { effect: 'fadeOut' };
    if (presetClass === 'emph') {
      if (presetID === 8 || presetID === 9) return { effect: 'spin' };
      if (presetID === 6 || presetID === 23) return { effect: 'zoom' };
      return { effect: 'pulse' };
    }
    const dirMap: Record<number, 'l' | 'r' | 'u' | 'd'> = {
      1: 'u',
      2: 'r',
      4: 'd',
      8: 'l',
      3: 'u',
      6: 'd',
      9: 'u',
      12: 'd',
    };
    const dir = dirMap[presetSubtype] || 'd';

    if (presetID === 1) return { effect: 'appear' };
    if (presetID === 2 || presetID === 3 || presetID === 12 || presetID === 21 || presetID === 22) {
      return { effect: 'fly', dir };
    }
    if (presetID === 10 || presetID === 9 || filterStr.includes('fade')) return { effect: 'fade' };
    if (presetID === 14 || presetID === 16 || presetID === 5 || filterStr.includes('wipe')) {
      return { effect: 'wipe', dir };
    }
    if (presetID === 23 || presetID === 53 || presetID === 54) return { effect: 'zoom' };
    if (presetID === 26 || presetID === 36) return { effect: 'bounce' };
    if (presetID === 42 || presetID === 43 || presetID === 45) return { effect: 'float', dir: 'd' };
    return { effect: 'fade' };
  };

  const mainSeqNodes = descEls(timing, 'cTn').filter(el => el.getAttribute('nodeType') === 'mainSeq');
  const roots = mainSeqNodes.length ? mainSeqNodes : [timing];

  const extractLeafAnimation = (parEl: Element, trigger: PptxAnimStep['trigger'], groupIdx: number, baseDelay: number) => {
    const cTn = firstChild(parEl, 'cTn');
    const spTargets = descEls(parEl, 'spTgt');
    const spid = spTargets.map(t => t.getAttribute('spid')).find(Boolean) || '';
    if (!spid) {
      const nestedChildLst = firstChild(cTn, 'childTnLst');
      if (nestedChildLst) {
        for (const subPar of childEls(nestedChildLst, 'par')) {
          extractLeafAnimation(subPar, trigger, groupIdx, baseDelay);
        }
      }
      return;
    }

    const allCtns = [cTn, ...descEls(parEl, 'cTn')].filter(Boolean) as Element[];
    const presetCtn = allCtns.find(x => x.getAttribute('presetClass') || x.getAttribute('presetID'));
    const rawClass = presetCtn?.getAttribute('presetClass') || 'entr';
    const presetClass: 'entr' | 'emph' | 'exit' | 'path' =
      rawClass === 'emph' ? 'emph' : rawClass === 'exit' ? 'exit' : rawClass === 'path' ? 'path' : 'entr';
    const presetID = parseInt(presetCtn?.getAttribute('presetID') || '10', 10) || 10;
    const presetSubtype = parseInt(presetCtn?.getAttribute('presetSubtype') || '4', 10) || 4;

    const animEffect = descEls(parEl, 'animEffect')[0];
    const filterStr = animEffect?.getAttribute('filter') || '';

    // Check for motion path (<p:animMotion path="M 0 0 L 0.25 0.1 E" />)
    const animMotion = descEls(parEl, 'animMotion')[0];
    let motionDxPct: number | undefined;
    let motionDyPct: number | undefined;
    if (animMotion) {
      const pathStr = animMotion.getAttribute('path') || '';
      const nums = pathStr.match(/-?\d+(\.\d+)?/g)?.map(Number);
      if (nums && nums.length >= 4) {
        motionDxPct = (nums[nums.length - 2] - nums[0]) * 100;
        motionDyPct = (nums[nums.length - 1] - nums[1]) * 100;
      }
    }

    const durCandidates = descEls(parEl, 'cTn')
      .map(x => x.getAttribute('dur'))
      .filter((v): v is string => !!v && v !== 'indefinite' && !Number.isNaN(Number(v)))
      .map(Number)
      .filter(n => n > 0 && n <= 8000);
    const durationMs = durCandidates.length ? Math.max(...durCandidates) : presetID === 1 ? 150 : 550;

    const stCond = firstChild(cTn, 'stCondLst');
    const condDelayAttr = stCond ? firstChild(stCond, 'cond')?.getAttribute('delay') : null;
    const ownDelay = condDelayAttr && condDelayAttr !== 'indefinite' ? Math.min(4000, Math.max(0, parseInt(condDelayAttr, 10) || 0)) : 0;

    const mapped = animMotion ? { effect: 'motionPath' as const } : mapPreset(presetClass, presetID, presetSubtype, filterStr);
    if (!steps.some(s => s.spid === spid && s.clickGroup === groupIdx && s.presetClass === presetClass)) {
      steps.push({
        spid,
        presetClass: animMotion ? 'path' : presetClass,
        effect: mapped.effect,
        dir: mapped.dir,
        motionDxPct,
        motionDyPct,
        trigger,
        durationMs,
        delayMs: baseDelay + ownDelay,
        clickGroup: groupIdx,
      });
    }
  };

  for (const root of roots) {
    const childLst = firstChild(root, 'childTnLst') || root;
    let runningDelay = 0;
    for (const par of childEls(childLst, 'par')) {
      const cTn = firstChild(par, 'cTn');
      const innerLst = firstChild(cTn, 'childTnLst');
      const subPars = innerLst ? childEls(innerLst, 'par') : [];
      const candidates = subPars.length ? subPars : [par];

      for (const item of candidates) {
        const itemCtn = firstChild(item, 'cTn');
        const nt = itemCtn?.getAttribute('nodeType') || 'clickEffect';
        if (nt === 'clickEffect') {
          clickGroup += 1;
          runningDelay = 0;
          extractLeafAnimation(item, 'onClick', clickGroup, 0);
        } else if (nt === 'afterEffect') {
          runningDelay += 250;
          extractLeafAnimation(item, 'afterPrev', clickGroup, runningDelay);
        } else {
          extractLeafAnimation(item, 'withPrev', clickGroup, runningDelay);
        }
      }
    }
  }

  return steps;
};

interface PlaceholderGeom {
  xEmu: number;
  yEmu: number;
  wEmu: number;
  hEmu: number;
  defaultFontPt?: number;
  fill?: string;
  textColor?: string;
}

const extractPlaceholderMap = (doc: Document | null, themeColors: Record<string, string>): Record<string, PlaceholderGeom> => {
  const map: Record<string, PlaceholderGeom> = {};
  if (!doc) return map;
  for (const sp of descEls(doc, 'sp')) {
    const nv = firstChild(sp, 'nvSpPr');
    const nvPr = firstChild(nv, 'nvPr');
    const ph = firstChild(nvPr, 'ph');
    if (!ph) continue;
    const type = ph.getAttribute('type') || 'body';
    const idx = ph.getAttribute('idx') || '';
    const spPr = firstChild(sp, 'spPr');
    const xfrm = firstChild(spPr, 'xfrm');
    const off = firstChild(xfrm, 'off');
    const ext = firstChild(xfrm, 'ext');
    if (!off || !ext) continue;
    const { fill } = extractFillStyle(spPr, themeColors);
    const txBody = firstChild(sp, 'txBody');
    const lstStyle = firstChild(txBody, 'lstStyle');
    const lvl1pPr = firstChild(lstStyle, 'lvl1pPr');
    const defRPr = firstChild(lvl1pPr, 'defRPr');
    const textColor = resolveColorEl(defRPr, themeColors);
    const szAttr = defRPr?.getAttribute('sz');
    const parsedFontPt = szAttr ? Math.max(10, Math.min(80, parseInt(szAttr, 10) / 100)) : undefined;

    const geom: PlaceholderGeom = {
      xEmu: parseInt(off.getAttribute('x') || '0', 10) || 0,
      yEmu: parseInt(off.getAttribute('y') || '0', 10) || 0,
      wEmu: parseInt(ext.getAttribute('cx') || '0', 10) || 0,
      hEmu: parseInt(ext.getAttribute('cy') || '0', 10) || 0,
      defaultFontPt: parsedFontPt || (type === 'title' || type === 'ctrTitle' ? 34 : type === 'subTitle' ? 20 : 18),
      fill,
      textColor,
    };
    if (type) map[`type:${type}`] = geom;
    if (idx) map[`idx:${idx}`] = geom;
  }
  return map;
};

let utifLoadingPromise: Promise<any> | null = null;
const loadUtif = (): Promise<any> => {
  if (typeof window === 'undefined') return Promise.resolve(null);
  if ((window as any).UTIF) return Promise.resolve((window as any).UTIF);
  if (utifLoadingPromise) return utifLoadingPromise;
  utifLoadingPromise = new Promise(resolve => {
    const s = document.createElement('script');
    s.src = 'https://cdn.jsdelivr.net/npm/utif@3.1.0/UTIF.min.js';
    s.onload = () => resolve((window as any).UTIF || null);
    s.onerror = () => resolve(null);
    document.head.appendChild(s);
  });
  return utifLoadingPromise;
};

const verifyImageBlobDecodes = (blob: Blob): Promise<Blob | null> =>
  new Promise(resolve => {
    if (!blob || blob.size === 0) return resolve(null);
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      URL.revokeObjectURL(url);
      resolve(blob);
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      resolve(null);
    };
    img.src = url;
  });

const extractEmbeddedRasterFromMetafile = (u8: Uint8Array): Blob | null => {
  // 1. Search for embedded PNG signature (89 50 4E 47 0D 0A 1A 0A ... 49 45 4E 44 AE 42 60 82)
  for (let i = 0; i <= u8.length - 16; i++) {
    if (
      u8[i] === 0x89 &&
      u8[i + 1] === 0x50 &&
      u8[i + 2] === 0x4e &&
      u8[i + 3] === 0x47 &&
      u8[i + 4] === 0x0d &&
      u8[i + 5] === 0x0a &&
      u8[i + 6] === 0x1a &&
      u8[i + 7] === 0x0a
    ) {
      for (let j = i + 16; j <= u8.length - 8; j++) {
        if (
          u8[j] === 0x49 &&
          u8[j + 1] === 0x45 &&
          u8[j + 2] === 0x4e &&
          u8[j + 3] === 0x44 &&
          u8[j + 4] === 0xae &&
          u8[j + 5] === 0x42 &&
          u8[j + 6] === 0x60 &&
          u8[j + 7] === 0x82
        ) {
          return new Blob([u8.slice(i, j + 8).buffer as ArrayBuffer], { type: 'image/png' });
        }
      }
    }
  }

  // 2. Search for embedded JPEG signature (FF D8 FF ... FF D9)
  for (let i = 0; i <= u8.length - 32; i++) {
    if (u8[i] === 0xff && u8[i + 1] === 0xd8 && u8[i + 2] === 0xff && (u8[i + 3] === 0xe0 || u8[i + 3] === 0xe1 || u8[i + 3] === 0xdb)) {
      let endIdx = -1;
      for (let j = u8.length - 2; j > i + 32; j--) {
        if (u8[j] === 0xff && u8[j + 1] === 0xd9) {
          endIdx = j + 2;
          break;
        }
      }
      if (endIdx > i) {
        return new Blob([u8.slice(i, endIdx).buffer as ArrayBuffer], { type: 'image/jpeg' });
      }
    }
  }

  // 3. Check EMF records for EMR_STRETCHDIBITS (81) or EMR_BITBLT (76) containing a DIB bitmap
  if (u8.byteLength >= 88) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    if (dv.getUint32(0, true) === 1 && dv.getUint32(40, true) === 0x464d4520) {
      let offset = 0;
      while (offset + 8 <= u8.byteLength) {
        const iType = dv.getUint32(offset, true);
        const nSize = dv.getUint32(offset + 4, true);
        if (nSize < 8 || offset + nSize > u8.byteLength) break;

        let offBmi = 0;
        let cbBmi = 0;
        let offBits = 0;
        let cbBits = 0;
        if (iType === 81 && nSize >= 80) {
          // EMR_STRETCHDIBITS
          offBmi = dv.getUint32(offset + 48, true);
          cbBmi = dv.getUint32(offset + 52, true);
          offBits = dv.getUint32(offset + 56, true);
          cbBits = dv.getUint32(offset + 60, true);
        } else if (iType === 76 && nSize >= 100) {
          // EMR_BITBLT
          offBmi = dv.getUint32(offset + 84, true);
          cbBmi = dv.getUint32(offset + 88, true);
          offBits = dv.getUint32(offset + 92, true);
          cbBits = dv.getUint32(offset + 96, true);
        }

        if (cbBmi >= 40 && cbBits > 0 && offset + offBmi + cbBmi <= u8.byteLength && offset + offBits + cbBits <= u8.byteLength) {
          const totalSize = 14 + cbBmi + cbBits;
          const bmpBuf = new Uint8Array(totalSize);
          const bmpView = new DataView(bmpBuf.buffer);
          bmpBuf[0] = 0x42; // 'B'
          bmpBuf[1] = 0x4d; // 'M'
          bmpView.setUint32(2, totalSize, true);
          bmpView.setUint32(6, 0, true);
          bmpView.setUint32(10, 14 + cbBmi, true);
          bmpBuf.set(u8.subarray(offset + offBmi, offset + offBmi + cbBmi), 14);
          bmpBuf.set(u8.subarray(offset + offBits, offset + offBits + cbBits), 14 + cbBmi);
          return new Blob([bmpBuf], { type: 'image/bmp' });
        }

        if (iType === 14) break; // EMR_EOF
        offset += nSize;
      }
    }
  }

  return null;
};

export const repairAndVerifyMediaBlob = async (
  input: Blob | ArrayBuffer,
  extHint = '',
  shrinkImage?: (ab: ArrayBuffer, ext: string) => Promise<Blob>
): Promise<Blob | null> => {
  try {
    const ab = input instanceof ArrayBuffer ? input : await input.arrayBuffer();
    if (!ab || ab.byteLength === 0) return null;
    const u8 = new Uint8Array(ab);

    // 1. Decompress gzip streams (.svgz, .emz, .wmz)
    if (u8.length >= 2 && u8[0] === 0x1f && u8[1] === 0x8b && typeof DecompressionStream !== 'undefined') {
      try {
        const ds = new DecompressionStream('gzip');
        const decompressedAb = await new Response(new Blob([ab]).stream().pipeThrough(ds)).arrayBuffer();
        const nextExt = extHint === 'svgz' ? 'svg' : extHint === 'emz' ? 'emf' : extHint === 'wmz' ? 'wmf' : extHint;
        return repairAndVerifyMediaBlob(decompressedAb, nextExt, shrinkImage);
      } catch {}
    }

    // 2. Magic byte detection for standard web raster formats
    if (u8.length >= 4 && u8[0] === 0x89 && u8[1] === 0x50 && u8[2] === 0x4e && u8[3] === 0x47) {
      const b = shrinkImage ? await shrinkImage(ab, 'png') : new Blob([ab], { type: 'image/png' });
      return verifyImageBlobDecodes(b);
    }
    if (u8.length >= 3 && u8[0] === 0xff && u8[1] === 0xd8 && u8[2] === 0xff) {
      const b = shrinkImage ? await shrinkImage(ab, 'jpg') : new Blob([ab], { type: 'image/jpeg' });
      return verifyImageBlobDecodes(b);
    }
    if (u8.length >= 4 && u8[0] === 0x47 && u8[1] === 0x49 && u8[2] === 0x46 && u8[3] === 0x38) {
      return verifyImageBlobDecodes(new Blob([ab], { type: 'image/gif' }));
    }
    if (
      u8.length >= 12 &&
      u8[0] === 0x52 &&
      u8[1] === 0x49 &&
      u8[2] === 0x46 &&
      u8[3] === 0x46 &&
      u8[8] === 0x57 &&
      u8[9] === 0x45 &&
      u8[10] === 0x42 &&
      u8[11] === 0x50
    ) {
      return verifyImageBlobDecodes(new Blob([ab], { type: 'image/webp' }));
    }
    if (u8.length >= 2 && u8[0] === 0x42 && u8[1] === 0x4d) {
      return verifyImageBlobDecodes(new Blob([ab], { type: 'image/bmp' }));
    }

    // 3. TIFF detection (II*\0 or MM\0*) -> convert to PNG using UTIF
    const isTiff =
      u8.length >= 4 &&
      ((u8[0] === 0x49 && u8[1] === 0x49 && u8[2] === 0x2a && u8[3] === 0x00) ||
        (u8[0] === 0x4d && u8[1] === 0x4d && u8[2] === 0x00 && u8[3] === 0x2a));
    if (isTiff || extHint === 'tif' || extHint === 'tiff') {
      const UTIF = await loadUtif();
      if (UTIF) {
        try {
          const ifds = UTIF.decode(ab);
          if (ifds && ifds.length > 0) {
            UTIF.decodeImage(ab, ifds[0]);
            const rgba = UTIF.toRGBA8(ifds[0]);
            const w = ifds[0].width;
            const h = ifds[0].height;
            if (w > 0 && h > 0 && rgba && rgba.length >= w * h * 4) {
              const canvas = document.createElement('canvas');
              canvas.width = w;
              canvas.height = h;
              const ctx = canvas.getContext('2d');
              if (ctx) {
                const clamped = new Uint8ClampedArray(w * h * 4);
                clamped.set(rgba.subarray(0, w * h * 4));
                ctx.putImageData(new ImageData(clamped, w, h), 0, 0);
                const pngBlob = await new Promise<Blob | null>(r => canvas.toBlob(b => r(b), 'image/png'));
                if (pngBlob) return verifyImageBlobDecodes(pngBlob);
              }
            }
          }
        } catch {}
      }
      // Fallback: check if browser natively decodes this TIFF
      return verifyImageBlobDecodes(new Blob([ab], { type: 'image/tiff' }));
    }

    // 4. SVG detection & xmlns repair
    const headText = new TextDecoder('utf-8', { fatal: false }).decode(u8.subarray(0, Math.min(u8.length, 512))).trim();
    if (extHint === 'svg' || headText.startsWith('<svg') || (headText.startsWith('<?xml') && headText.includes('<svg')) || headText.includes('<svg')) {
      let fullSvg = new TextDecoder('utf-8', { fatal: false }).decode(u8);
      if (fullSvg.includes('<svg') && !fullSvg.includes('xmlns=')) {
        fullSvg = fullSvg.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"');
      }
      const svgBlob = new Blob([fullSvg], { type: 'image/svg+xml;charset=utf-8' });
      return verifyImageBlobDecodes(svgBlob);
    }

    // 5. EMF / WMF metafile: extract embedded PNG, JPEG, or DIB bitmap
    const extracted = extractEmbeddedRasterFromMetafile(u8);
    if (extracted) {
      return verifyImageBlobDecodes(extracted);
    }

    return null;
  } catch {
    return null;
  }
};

export const parsePptxFile = async (
  zip: any,
  putBlob: (b: Blob) => string,
  shrinkImage: (ab: ArrayBuffer, ext: string) => Promise<Blob>,
  onProgress?: (i: number, total: number) => void
): Promise<Array<{ title: string; body: string; imgs: string[]; vid?: string; tx: string; pptx: PptxSlideData }>> => {
  const themeColors = await parseThemeColors(zip);

  // Read slide size and exact slide order from ppt/presentation.xml
  let slideW = 12192000;
  let slideH = 6858000;
  let orderedSlidePaths: string[] = [];

  const presFile = zip.file('ppt/presentation.xml');
  if (presFile) {
    const presDoc = parseXmlDoc(await presFile.async('string'));
    const sldSz = descEls(presDoc, 'sldSz')[0];
    if (sldSz) {
      slideW = parseInt(sldSz.getAttribute('cx') || String(slideW), 10) || slideW;
      slideH = parseInt(sldSz.getAttribute('cy') || String(slideH), 10) || slideH;
    }
    const presRels = await parseRelsMap(zip, 'ppt/_rels/presentation.xml.rels', 'ppt');
    const sldIds = descEls(presDoc, 'sldId');
    for (const sldId of sldIds) {
      const rId = getRelAttr(sldId, 'id');
      const target = rId ? presRels[rId]?.target : '';
      if (target && zip.file(target)) {
        orderedSlidePaths.push(target);
      }
    }
  }

  const widthPx = 960;
  const heightPx = Math.max(360, Math.round(widthPx * (slideH / slideW)));

  // Cache media blobs so shared images/videos/audio across slides aren't duplicated
  const mediaBlobCache: Record<string, { id: string; kind: 'img' | 'vid' | 'aud' } | null> = {};

  const findZipEntry = (zipPath: string) => {
    if (!zipPath) return null;
    const direct = zip.file(zipPath);
    if (direct) return direct;
    const lower = zipPath.toLowerCase();
    const matchedKey = Object.keys(zip.files).find(k => k.toLowerCase() === lower);
    return matchedKey ? zip.file(matchedKey) : null;
  };

  const loadMediaByPath = async (zipPath: string): Promise<{ id: string; kind: 'img' | 'vid' | 'aud' } | null> => {
    if (!zipPath || /^https?:\/\//i.test(zipPath)) return null;
    if (zipPath in mediaBlobCache) return mediaBlobCache[zipPath];
    const zf = findZipEntry(zipPath);
    if (!zf) {
      mediaBlobCache[zipPath] = null;
      return null;
    }
    const ext = zipPath.split('.').pop()?.toLowerCase() || '';

    if (['mp4', 'm4v', 'webm', 'mov', 'ogv', 'avi', 'wmv', 'mkv'].includes(ext)) {
      const ab = await zf.async('arraybuffer');
      if (!ab || ab.byteLength === 0) {
        mediaBlobCache[zipPath] = null;
        return null;
      }
      const mimeMap: Record<string, string> = {
        mp4: 'video/mp4',
        m4v: 'video/mp4',
        webm: 'video/webm',
        mov: 'video/mp4',
        ogv: 'video/ogg',
        avi: 'video/x-msvideo',
        wmv: 'video/x-ms-wmv',
        mkv: 'video/x-matroska',
      };
      const vb = new Blob([ab], { type: mimeMap[ext] || 'video/mp4' });
      const entry = { id: putBlob(vb), kind: 'vid' as const };
      mediaBlobCache[zipPath] = entry;
      return entry;
    }

    if (['mp3', 'wav', 'm4a', 'ogg', 'aac', 'wma', 'flac'].includes(ext)) {
      const ab = await zf.async('arraybuffer');
      if (!ab || ab.byteLength === 0) {
        mediaBlobCache[zipPath] = null;
        return null;
      }
      const mimeMap: Record<string, string> = {
        mp3: 'audio/mpeg',
        wav: 'audio/wav',
        m4a: 'audio/mp4',
        ogg: 'audio/ogg',
        aac: 'audio/aac',
        wma: 'audio/x-ms-wma',
        flac: 'audio/flac',
      };
      const ablob = new Blob([ab], { type: mimeMap[ext] || 'audio/mpeg' });
      const entry = { id: putBlob(ablob), kind: 'aud' as const };
      mediaBlobCache[zipPath] = entry;
      return entry;
    }

    // Handle all image and metafile formats (.png, .jpg, .jpeg, .gif, .webp, .svg, .svgz, .bmp, .jfif, .tif, .tiff, .emf, .wmf, .emz, .wmz)
    if (['png', 'jpg', 'jpeg', 'gif', 'webp', 'svg', 'svgz', 'bmp', 'jfif', 'tif', 'tiff', 'emf', 'wmf', 'emz', 'wmz', 'wdp', 'hdp', 'bin', 'tmp'].includes(ext)) {
      const ab = await zf.async('arraybuffer');
      const repaired = await repairAndVerifyMediaBlob(ab, ext, shrinkImage);
      if (!repaired) {
        mediaBlobCache[zipPath] = null;
        return null;
      }
      const entry = { id: putBlob(repaired), kind: 'img' as const };
      mediaBlobCache[zipPath] = entry;
      return entry;
    }

    mediaBlobCache[zipPath] = null;
    return null;
  };

  const fallbackSlideKeys = Object.keys(zip.files)
    .filter(k => /^ppt\/slides\/slide\d+\.xml$/.test(k))
    .sort((a, b) => parseInt(a.match(/\d+/)![0], 10) - parseInt(b.match(/\d+/)![0], 10));

  const slideKeys = orderedSlidePaths.length ? orderedSlidePaths : fallbackSlideKeys;
  const out: Array<{ title: string; body: string; imgs: string[]; vid?: string; tx: string; pptx: PptxSlideData }> = [];

  for (const [i, k] of slideKeys.entries()) {
    onProgress?.(i, slideKeys.length);
    const slideText = await zip.file(k)?.async('string') || '';
    const pXml = parseXmlDoc(slideText);
    const slideRelsPath = k.replace('slides/', 'slides/_rels/') + '.rels';
    const slideRels = await parseRelsMap(zip, slideRelsPath, 'ppt/slides');

    let layoutDoc: Document | null = null;
    let layoutRels: Record<string, { target: string; type: string }> = {};
    let masterDoc: Document | null = null;
    let masterRels: Record<string, { target: string; type: string }> = {};

    const layoutRel = Object.values(slideRels).find(r => r.type.endsWith('/slideLayout') || r.target.includes('slideLayout'));
    if (layoutRel && zip.file(layoutRel.target)) {
      layoutDoc = parseXmlDoc(await zip.file(layoutRel.target).async('string'));
      const layoutName = layoutRel.target.split('/').pop()!;
      const layoutBase = layoutRel.target.slice(0, layoutRel.target.lastIndexOf('/'));
      layoutRels = await parseRelsMap(zip, `${layoutBase}/_rels/${layoutName}.rels`, layoutBase);

      const masterRel = Object.values(layoutRels).find(r => r.type.endsWith('/slideMaster') || r.target.includes('slideMaster'));
      if (masterRel && zip.file(masterRel.target)) {
        masterDoc = parseXmlDoc(await zip.file(masterRel.target).async('string'));
        const masterName = masterRel.target.split('/').pop()!;
        const masterBase = masterRel.target.slice(0, masterRel.target.lastIndexOf('/'));
        masterRels = await parseRelsMap(zip, `${masterBase}/_rels/${masterName}.rels`, masterBase);
      }
    }

    const phMap = {
      ...extractPlaceholderMap(masterDoc, themeColors),
      ...extractPlaceholderMap(layoutDoc, themeColors),
    };

    const slideMediaIds: string[] = [];
    const slideAudioIds: string[] = [];
    let slideVidId: string | undefined;

    const registerMedia = (m: { id: string; kind: 'img' | 'vid' | 'aud' } | null) => {
      if (!m?.id) return;
      if (!slideMediaIds.includes(m.id)) slideMediaIds.push(m.id);
      if (m.kind === 'vid' && !slideVidId) slideVidId = m.id;
      if (m.kind === 'aud' && !slideAudioIds.includes(m.id)) slideAudioIds.push(m.id);
    };

    // Extract background from slide -> layout -> master
    let background: string | undefined;
    let bgImgId: string | undefined;
    const bgSources: Array<{ doc: Document | null; rels: Record<string, { target: string; type: string }> }> = [
      { doc: pXml, rels: slideRels },
      { doc: layoutDoc, rels: layoutRels },
      { doc: masterDoc, rels: masterRels },
    ];
    for (const src of bgSources) {
      if (!src.doc || background || bgImgId) continue;
      const bgEl = descEls(src.doc, 'bg')[0];
      if (!bgEl) continue;
      const bgPr = firstChild(bgEl, 'bgPr');
      if (bgPr) {
        const blipFill = firstChild(bgPr, 'blipFill');
        const blip = blipFill ? firstChild(blipFill, 'blip') : null;
        const svgBlip = blip ? descEls(blip, 'svgBlip')[0] : null;
        const embedIds = [
          getRelAttr(blip, 'embed'),
          getRelAttr(blip, 'link'),
          getRelAttr(svgBlip, 'embed'),
          getRelAttr(svgBlip, 'link'),
        ].filter(Boolean);
        for (const eid of embedIds) {
          if (src.rels[eid]) {
            const media = await loadMediaByPath(src.rels[eid].target);
            if (media?.id && media.kind === 'img') {
              bgImgId = media.id;
              registerMedia(media);
              break;
            }
          }
        }
        const f = extractFillStyle(bgPr, themeColors);
        if (f.fill) background = f.fill;
      }
      const bgRef = firstChild(bgEl, 'bgRef');
      if (bgRef && !background) {
        const c = resolveColorEl(bgRef, themeColors);
        if (c) background = c;
      }
    }

    const slideIsDark = isDarkColor(background);
    const elements: PptxElement[] = [];

    const clampBoxPct = (box: { x: number; y: number; w: number; h: number }) => {
      const rawXPct = (box.x / slideW) * 100;
      const rawYPct = (box.y / slideH) * 100;
      const rawWPct = (box.w / slideW) * 100;
      const rawHPct = (box.h / slideH) * 100;
      const safeXPct = Math.max(0, Math.min(98, rawXPct));
      const safeYPct = Math.max(0, Math.min(98, rawYPct));
      const safeWPct = Math.max(0.5, Math.min(100 - safeXPct, rawWPct));
      const safeHPct = Math.max(0.5, Math.min(100 - safeYPct, rawHPct));
      return { xPct: safeXPct, yPct: safeYPct, wPct: safeWPct, hPct: safeHPct };
    };

    const parseXfrm = (
      xfrmEl: Element | null | undefined,
      parentScale?: { ox: number; oy: number; sx: number; sy: number }
    ) => {
      if (!xfrmEl) return null;
      const off = firstChild(xfrmEl, 'off');
      const ext = firstChild(xfrmEl, 'ext');
      if (!off || !ext) return null;
      let x = parseInt(off.getAttribute('x') || '0', 10) || 0;
      let y = parseInt(off.getAttribute('y') || '0', 10) || 0;
      let w = parseInt(ext.getAttribute('cx') || '0', 10) || 0;
      let h = parseInt(ext.getAttribute('cy') || '0', 10) || 0;
      if (parentScale) {
        x = parentScale.ox + x * parentScale.sx;
        y = parentScale.oy + y * parentScale.sy;
        w = w * parentScale.sx;
        h = h * parentScale.sy;
      }
      const rot = xfrmEl.getAttribute('rot');
      const rotDeg = rot ? Math.round((parseInt(rot, 10) || 0) / 60000) : 0;
      const flipH = xfrmEl.getAttribute('flipH') === '1';
      const flipV = xfrmEl.getAttribute('flipV') === '1';
      return { x, y, w, h, rotDeg, flipH, flipV };
    };

    // Extract all image, video, and audio references inside a shape/pic/frame node
    const extractNodeMedia = async (
      node: Element,
      rels: Record<string, { target: string; type: string }>
    ): Promise<{ imgId?: string; vidId?: string; audioId?: string }> => {
      let imgId: string | undefined;
      let vidId: string | undefined;
      let audioId: string | undefined;

      // Check video/audio references (videoFile, audioFile, quickTimeFile, p14:media)
      const mediaTags = [
        ...descEls(node, 'videoFile'),
        ...descEls(node, 'audioFile'),
        ...descEls(node, 'quickTimeFile'),
        ...descEls(node, 'media'),
        ...descEls(node, 'wavAudioFile'),
      ];
      for (const mEl of mediaTags) {
        const rIds = [getRelAttr(mEl, 'embed'), getRelAttr(mEl, 'link')].filter(Boolean);
        for (const rId of rIds) {
          if (rels[rId]) {
            const loaded = await loadMediaByPath(rels[rId].target);
            if (loaded?.id) {
              registerMedia(loaded);
              if (loaded.kind === 'vid' && !vidId) vidId = loaded.id;
              if (loaded.kind === 'aud' && !audioId) audioId = loaded.id;
              if (loaded.kind === 'img' && !imgId) imgId = loaded.id;
            }
          }
        }
      }

      // Check blips (prefer standard raster blip r:embed first, then svgBlip as fallback)
      const blips = descEls(node, 'blip');
      for (const blip of blips) {
        if (imgId) break;
        const svgBlip = descEls(blip, 'svgBlip')[0];
        const rIds = [
          getRelAttr(blip, 'embed'),
          getRelAttr(blip, 'link'),
          getRelAttr(svgBlip, 'embed'),
          getRelAttr(svgBlip, 'link'),
        ].filter(Boolean);
        for (const rId of rIds) {
          if (rels[rId]) {
            const loaded = await loadMediaByPath(rels[rId].target);
            if (loaded?.id) {
              registerMedia(loaded);
              if (loaded.kind === 'img' && !imgId) {
                imgId = loaded.id;
                break;
              }
              if (loaded.kind === 'vid' && !vidId) vidId = loaded.id;
            }
          }
        }
      }

      // Check VML imagedata (<v:imagedata r:id="rId..." o:relid="rId..." />) for legacy OLE/equations
      if (!imgId) {
        const vmlImgs = descEls(node, 'imagedata');
        for (const vImg of vmlImgs) {
          if (imgId) break;
          const rIds = [
            getRelAttr(vImg, 'id'),
            getRelAttr(vImg, 'href'),
            vImg.getAttribute('o:relid') || '',
            vImg.getAttribute('r:id') || '',
          ].filter(Boolean);
          for (const rId of rIds) {
            if (rels[rId]) {
              const loaded = await loadMediaByPath(rels[rId].target);
              if (loaded?.id && loaded.kind === 'img') {
                registerMedia(loaded);
                imgId = loaded.id;
                break;
              }
            }
          }
        }
      }

      return { imgId, vidId, audioId };
    };

    const parseParagraphsFromTxBody = (
      txBody: Element | null,
      defaultFontPt: number,
      defaultAlign: PptxParagraph['align'],
      fallbackTextColor: string
    ): PptxParagraph[] => {
      const paragraphs: PptxParagraph[] = [];
      if (!txBody) return paragraphs;

      for (const pEl of childEls(txBody, 'p')) {
        const pPr = firstChild(pEl, 'pPr');
        const algnAttr = pPr?.getAttribute('algn');
        const align: PptxParagraph['align'] =
          algnAttr === 'ctr'
            ? 'center'
            : algnAttr === 'r'
            ? 'right'
            : algnAttr === 'just'
            ? 'justify'
            : algnAttr === 'l'
            ? 'left'
            : defaultAlign;
        const hasBullet = !!(pPr && !firstChild(pPr, 'buNone') && (firstChild(pPr, 'buChar') || firstChild(pPr, 'buAutoNum')));
        const pDefRPr = firstChild(pPr, 'defRPr');

        const runs: PptxTextRun[] = [];
        for (const pChild of childEls(pEl)) {
          if (pChild.localName === 'r' || pChild.localName === 'fld') {
            const rPr = firstChild(pChild, 'rPr');
            const tEl = firstChild(pChild, 't');
            const txt = tEl?.textContent || '';
            if (!txt) continue;
            const szAttr = rPr?.getAttribute('sz') || pDefRPr?.getAttribute('sz');
            const fontSizePt = szAttr ? Math.max(8, Math.min(96, parseInt(szAttr, 10) / 100)) : defaultFontPt;
            const runColor =
              resolveColorEl(rPr, themeColors) ||
              resolveColorEl(pDefRPr, themeColors) ||
              fallbackTextColor;
            runs.push({
              text: txt,
              bold: rPr?.getAttribute('b') === '1' || pDefRPr?.getAttribute('b') === '1',
              italic: rPr?.getAttribute('i') === '1' || pDefRPr?.getAttribute('i') === '1',
              underline: !!rPr?.getAttribute('u') && rPr.getAttribute('u') !== 'none',
              fontSizePt,
              color: runColor,
            });
          } else if (pChild.localName === 'br') {
            runs.push({ text: '\n' });
          } else if (pChild.localName === 'm') {
            // Office Math (<m:oMath>) text fallback
            const mathTxt = descEls(pChild, 't').map(t => t.textContent || '').join('');
            if (mathTxt) {
              runs.push({ text: mathTxt, italic: true, fontSizePt: defaultFontPt, color: fallbackTextColor });
            }
          }
        }

        if (runs.some(r => r.text.trim())) {
          paragraphs.push({ align, bullet: hasBullet, runs });
        }
      }
      return paragraphs;
    };

    // Unwrap <mc:AlternateContent> inside spTree or grpSp so videos/SVGs/shapes/OLE previews are properly selected
    const expandTreeChildren = (treeEl: Element): Element[] => {
      const result: Element[] = [];
      for (const c of childEls(treeEl)) {
        if (c.localName === 'AlternateContent') {
          const choice = firstChild(c, 'Choice');
          const fallback = firstChild(c, 'Fallback');
          const choiceChildren = choice ? childEls(choice) : [];
          const fallbackChildren = fallback ? childEls(fallback) : [];
          const choiceHasAVMedia =
            choice &&
            (descEls(choice, 'videoFile').length > 0 ||
              descEls(choice, 'audioFile').length > 0 ||
              descEls(choice, 'media').length > 0 ||
              descEls(choice, 'quickTimeFile').length > 0);
          const choiceIsProprietary =
            choice &&
            !choiceHasAVMedia &&
            (descEls(choice, 'oleObj').length > 0 ||
              descEls(choice, 'contentPart').length > 0 ||
              choiceChildren.every(
                ch => ch.localName !== 'pic' && ch.localName !== 'sp' && ch.localName !== 'grpSp' && ch.localName !== 'graphicFrame'
              ));

          if (choiceIsProprietary && fallbackChildren.length > 0) {
            result.push(...fallbackChildren);
          } else if (choiceChildren.length > 0) {
            result.push(...choiceChildren);
          } else if (fallbackChildren.length > 0) {
            result.push(...fallbackChildren);
          }
        } else if (c.localName === 'oleObj') {
          const innerPics = childEls(c, 'pic');
          if (innerPics.length > 0) result.push(...innerPics);
          else result.push(c);
        } else if (c.localName === 'contentPart') {
          // Handled via rels if needed
        } else {
          result.push(c);
        }
      }
      return result;
    };

    const walkSpTree = async (
      treeEl: Element | null,
      rels: Record<string, { target: string; type: string }>,
      onlyNonPlaceholders = false,
      parentScale?: { ox: number; oy: number; sx: number; sy: number }
    ) => {
      if (!treeEl) return;
      for (const child of expandTreeChildren(treeEl)) {
        if (child.localName === 'grpSp') {
          const grpPr = firstChild(child, 'grpSpPr');
          const xfrm = firstChild(grpPr, 'xfrm');
          const off = firstChild(xfrm, 'off');
          const ext = firstChild(xfrm, 'ext');
          const chOff = firstChild(xfrm, 'chOff');
          const chExt = firstChild(xfrm, 'chExt');
          let nextScale = parentScale;
          if (off && ext && chOff && chExt) {
            const gx = parseInt(off.getAttribute('x') || '0', 10) || 0;
            const gy = parseInt(off.getAttribute('y') || '0', 10) || 0;
            const gw = parseInt(ext.getAttribute('cx') || '1', 10) || 1;
            const gh = parseInt(ext.getAttribute('cy') || '1', 10) || 1;
            const cx = parseInt(chOff.getAttribute('x') || '0', 10) || 0;
            const cy = parseInt(chOff.getAttribute('y') || '0', 10) || 0;
            const cw = parseInt(chExt.getAttribute('cx') || String(gw), 10) || gw;
            const ch = parseInt(chExt.getAttribute('cy') || String(gh), 10) || gh;
            const sx = gw / cw;
            const sy = gh / ch;
            const ox = gx - cx * sx;
            const oy = gy - cy * sy;
            nextScale = parentScale
              ? { ox: parentScale.ox + ox * parentScale.sx, oy: parentScale.oy + oy * parentScale.sy, sx: parentScale.sx * sx, sy: parentScale.sy * sy }
              : { ox, oy, sx, sy };
          }
          await walkSpTree(child, rels, onlyNonPlaceholders, nextScale);
        } else if (child.localName === 'pic') {
          const nvPicPr = firstChild(child, 'nvPicPr');
          const cNvPr = firstChild(nvPicPr, 'cNvPr');
          const nvPr = firstChild(nvPicPr, 'nvPr');
          if (onlyNonPlaceholders && firstChild(nvPr, 'ph')) continue;

          const spid = cNvPr?.getAttribute('id') || `pic-${elements.length}`;
          const spPr = firstChild(child, 'spPr');
          const box = parseXfrm(firstChild(spPr, 'xfrm'), parentScale);
          const { imgId, vidId, audioId } = await extractNodeMedia(child, rels);

          if (box && (imgId || vidId || audioId) && box.w > 0 && box.h > 0) {
            const coords = clampBoxPct(box);
            elements.push({
              id: spid,
              kind: 'pic',
              ...coords,
              rotDeg: box.rotDeg || undefined,
              flipH: box.flipH || undefined,
              flipV: box.flipV || undefined,
              imgId,
              vidId,
              audioId,
            });
          }
        } else if (child.localName === 'sp' || child.localName === 'cxnSp') {
          const nvSpPr = firstChild(child, 'nvSpPr') || firstChild(child, 'nvCxnSpPr');
          const cNvPr = firstChild(nvSpPr, 'cNvPr');
          const spid = cNvPr?.getAttribute('id') || `sp-${elements.length}`;
          const nvPr = firstChild(nvSpPr, 'nvPr');
          const ph = firstChild(nvPr, 'ph');
          if (onlyNonPlaceholders && ph) continue;

          const phType = ph?.getAttribute('type') || (ph ? 'body' : '');
          const phIdx = ph?.getAttribute('idx') || '';
          const fallbackPh =
            (phIdx && phMap[`idx:${phIdx}`]) ||
            (phType && phMap[`type:${phType}`]) ||
            (ph ? phMap['type:body'] || phMap['type:obj'] || phMap['idx:1'] : undefined);

          const spPr = firstChild(child, 'spPr');
          let box = parseXfrm(firstChild(spPr, 'xfrm'), parentScale);
          if (!box && fallbackPh && fallbackPh.wEmu > 0 && fallbackPh.hEmu > 0) {
            box = { x: fallbackPh.xEmu, y: fallbackPh.yEmu, w: fallbackPh.wEmu, h: fallbackPh.hEmu, rotDeg: 0, flipH: false, flipV: false };
          }
          if (!box || (box.w <= 0 && box.h <= 0)) continue;
          if (box.w <= 0) box.w = 12700;
          if (box.h <= 0) box.h = 12700;

          const prstGeom = firstChild(spPr, 'prstGeom');
          const geom = prstGeom?.getAttribute('prst') || 'rect';

          // Shape fill (explicit spPr, p:style/fillRef, or inherited placeholder fill)
          let { fill, noFill } = extractFillStyle(spPr, themeColors);
          if (!fill && !noFill) {
            const styleEl = firstChild(child, 'style');
            const fillRef = firstChild(styleEl, 'fillRef');
            if (fillRef && fillRef.getAttribute('idx') !== '0') {
              fill = resolveColorEl(fillRef, themeColors);
            } else if (fallbackPh?.fill) {
              fill = fallbackPh.fill;
            }
          }

          // Also check if shape has a picture fill (spPr > blipFill) or video/audio attached
          const { imgId: shapeImgId, vidId: shapeVidId, audioId: shapeAudioId } = await extractNodeMedia(child, rels);

          // Shape border
          const ln = firstChild(spPr, 'ln');
          let stroke: string | undefined;
          let strokeWidthPx: number | undefined;
          if (ln && !firstChild(ln, 'noFill')) {
            const lf = extractFillStyle(ln, themeColors);
            if (lf.fill) {
              stroke = lf.fill;
              const wEmu = parseInt(ln.getAttribute('w') || '12700', 10) || 12700;
              strokeWidthPx = Math.max(1, Math.round(wEmu / 12700));
            }
          }

          // Text body & runs
          const txBody = firstChild(child, 'txBody');
          const bodyPr = firstChild(txBody, 'bodyPr');
          const anchor = bodyPr?.getAttribute('anchor');
          const valign: 'top' | 'middle' | 'bottom' =
            anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top';

          const normAutofit = firstChild(bodyPr, 'normAutofit');
          const fontScaleAttr = normAutofit?.getAttribute('fontScale');
          const fontScale = fontScaleAttr
            ? Math.max(0.45, Math.min(1, (parseInt(fontScaleAttr, 10) || 100000) / 100000))
            : undefined;

          const shapeBgDark = fill ? isDarkColor(fill) : slideIsDark;
          const defaultFontPt =
            fallbackPh?.defaultFontPt ||
            (phType === 'title' || phType === 'ctrTitle' ? 34 : phType === 'subTitle' ? 20 : 18);
          const defaultAlign: PptxParagraph['align'] =
            phType === 'ctrTitle' || phType === 'subTitle' ? 'center' : 'left';
          const fallbackTextColor = fallbackPh?.textColor || (shapeBgDark ? '#f8fafc' : '#0f172a');

          const paragraphs = parseParagraphsFromTxBody(txBody, defaultFontPt, defaultAlign, fallbackTextColor);
          if (phType === 'title' || phType === 'ctrTitle') {
            paragraphs.forEach(p => p.runs.forEach(r => { if (r.bold === undefined) r.bold = true; }));
          }

          if (fill || stroke || paragraphs.length > 0 || shapeImgId || shapeVidId || shapeAudioId) {
            const coords = clampBoxPct(box);
            elements.push({
              id: spid,
              kind: (shapeImgId || shapeVidId) && paragraphs.length === 0 ? 'pic' : 'shape',
              ...coords,
              rotDeg: box.rotDeg || undefined,
              flipH: box.flipH || undefined,
              flipV: box.flipV || undefined,
              geom,
              fill,
              stroke,
              strokeWidthPx,
              paragraphs,
              valign,
              fontScale,
              imgId: shapeImgId,
              vidId: shapeVidId,
              audioId: shapeAudioId,
            });
          }
        } else if (child.localName === 'graphicFrame') {
          const nvGraphicFramePr = firstChild(child, 'nvGraphicFramePr');
          const cNvPr = firstChild(nvGraphicFramePr, 'cNvPr');
          const spid = cNvPr?.getAttribute('id') || `gf-${elements.length}`;
          const xfrm = firstChild(child, 'xfrm');
          const box = parseXfrm(xfrm, parentScale);

          // 1. Check if graphicFrame is a Table (<a:tbl>)
          const tbl = descEls(child, 'tbl')[0];
          if (tbl && box && box.w > 0 && box.h > 0) {
            const tableRows: PptxTableCell[][] = [];
            for (const tr of childEls(tbl, 'tr')) {
              const rowCells: PptxTableCell[] = [];
              for (const tc of childEls(tr, 'tc')) {
                if (tc.getAttribute('hMerge') === '1' || tc.getAttribute('vMerge') === '1') continue;
                const tcPr = firstChild(tc, 'tcPr');
                const { fill: cellFill } = extractFillStyle(tcPr, themeColors);
                const cellDark = cellFill ? isDarkColor(cellFill) : slideIsDark;
                const txBody = firstChild(tc, 'txBody');
                const cellParagraphs = parseParagraphsFromTxBody(txBody, 14, 'left', cellDark ? '#f8fafc' : '#0f172a');
                const cellText = cellParagraphs.map(p => p.runs.map(r => r.text).join('')).join('\n');
                const gridSpan = parseInt(tc.getAttribute('gridSpan') || '1', 10) || 1;
                const rowSpan = parseInt(tc.getAttribute('rowSpan') || '1', 10) || 1;
                rowCells.push({
                  text: cellText,
                  paragraphs: cellParagraphs,
                  fill: cellFill,
                  colSpan: gridSpan > 1 ? gridSpan : undefined,
                  rowSpan: rowSpan > 1 ? rowSpan : undefined,
                });
              }
              if (rowCells.length > 0) tableRows.push(rowCells);
            }
            if (tableRows.length > 0) {
              const coords = clampBoxPct(box);
              elements.push({
                id: spid,
                kind: 'table',
                ...coords,
                tableRows,
              });
              continue;
            }
          }

          // 2. Check if graphicFrame has media (video/audio/OLE preview image)
          const { imgId, vidId, audioId } = await extractNodeMedia(child, rels);
          if (box && (imgId || vidId || audioId) && box.w > 0 && box.h > 0) {
            const coords = clampBoxPct(box);
            elements.push({
              id: spid,
              kind: 'pic',
              ...coords,
              imgId,
              vidId,
              audioId,
            });
          }
        }
      }
    };

    // 1. Walk non-placeholder decorative background elements from slideMaster and slideLayout
    const slideRoot = pXml.documentElement;
    const layoutRoot = layoutDoc?.documentElement;
    const showMasterOnLayout = layoutRoot?.getAttribute('showMasterSp') !== '0';
    const showMasterOnSlide = slideRoot?.getAttribute('showMasterSp') !== '0';

    if (showMasterOnSlide && showMasterOnLayout && masterDoc) {
      const masterCSld = firstChild(masterDoc.documentElement, 'cSld');
      await walkSpTree(firstChild(masterCSld, 'spTree'), masterRels, true);
    }
    if (showMasterOnSlide && layoutDoc) {
      const layoutCSld = firstChild(layoutDoc.documentElement, 'cSld');
      await walkSpTree(firstChild(layoutCSld, 'spTree'), layoutRels, true);
    }

    // 2. Walk the slide's own spTree
    const cSld = firstChild(pXml.documentElement, 'cSld');
    const spTree = firstChild(cSld, 'spTree');
    await walkSpTree(spTree, slideRels, false);

    // 3. Walk any SmartArt / Diagram drawing XML referenced by this slide (ppt/diagrams/drawing*.xml)
    for (const rel of Object.values(slideRels)) {
      if (rel.target.includes('diagrams/drawing') && zip.file(rel.target)) {
        const dgmDoc = parseXmlDoc(await zip.file(rel.target).async('string'));
        const dgmName = rel.target.split('/').pop()!;
        const dgmBase = rel.target.slice(0, rel.target.lastIndexOf('/'));
        const dgmRels = await parseRelsMap(zip, `${dgmBase}/_rels/${dgmName}.rels`, dgmBase);
        const dgmSpTree = descEls(dgmDoc, 'spTree')[0];
        if (dgmSpTree) {
          await walkSpTree(dgmSpTree, { ...slideRels, ...dgmRels }, false);
        }
      }
    }

    // 4. Fallback: only if no visual elements or background image were parsed on the slide, check slideRels for standalone video/image
    const hasVisualElements = Boolean(bgImgId) || elements.some(el => Boolean(el.imgId || el.vidId || el.paragraphs?.length || el.tableRows?.length));
    for (const rel of Object.values(slideRels)) {
      if (!rel.target.includes('media/')) continue;
      // Always capture standalone videos or audio in slideRels even if slide has text shapes
      const ext = rel.target.split('.').pop()?.toLowerCase() || '';
      const isAV = ['mp4', 'm4v', 'webm', 'mov', 'ogv', 'avi', 'wmv', 'mkv', 'mp3', 'wav', 'm4a', 'ogg', 'aac', 'wma', 'flac'].includes(ext);
      if (hasVisualElements && !isAV) continue;

      const m = await loadMediaByPath(rel.target);
      if (!m?.id) continue;
      registerMedia(m);

      const alreadyPlaced =
        bgImgId === m.id ||
        elements.some(el => el.imgId === m.id || el.vidId === m.id || el.audioId === m.id);
      if (!alreadyPlaced) {
        if (m.kind === 'vid') {
          elements.push({
            id: `unplaced-vid-${elements.length}`,
            kind: 'pic',
            xPct: 10,
            yPct: 14,
            wPct: 80,
            hPct: 72,
            vidId: m.id,
          });
        } else if (m.kind === 'img' && !hasVisualElements) {
          elements.push({
            id: `unplaced-img-${elements.length}`,
            kind: 'pic',
            xPct: 12,
            yPct: 15,
            wPct: 76,
            hPct: 70,
            imgId: m.id,
          });
        }
      }
    }

    const transition =
      parseTransition(pXml) ||
      (layoutDoc ? parseTransition(layoutDoc) : undefined) ||
      (masterDoc ? parseTransition(masterDoc) : undefined);
    const animations = parseTimingAnimations(pXml);

    const textLines = descEls(pXml, 'p')
      .map(p => descEls(p, 't').map(t => t.textContent || '').join(''))
      .filter(x => x.trim());

    out.push({
      title: textLines[0] || `Slide ${i + 1}`,
      body: textLines.slice(1).join('\n'),
      imgs: slideMediaIds,
      vid: slideVidId,
      tx: textLines.join('\n'),
      pptx: {
        widthPx,
        heightPx,
        background: background || '#ffffff',
        bgImgId,
        elements,
        transition,
        animations,
        audioIds: slideAudioIds.length ? slideAudioIds : undefined,
      },
    });
  }

  return out;
};

// Transition motion variants
const getSlideTransitionProps = (transition?: PptxTransition) => {
  if (!transition || transition.type === 'none' || transition.type === 'cut') {
    return {
      initial: { opacity: 0.99 },
      animate: { opacity: 1 },
      exit: { opacity: 0 },
      transition: { duration: 0.12 },
    };
  }
  const dur = Math.max(0.2, Math.min(2, transition.durationMs / 1000));
  const dir = transition.dir || 'l';
  if (transition.type === 'push' || transition.type === 'cover') {
    const dx = dir === 'l' ? '18%' : dir === 'r' ? '-18%' : '0%';
    const dy = dir === 'u' ? '18%' : dir === 'd' ? '-18%' : '0%';
    return {
      initial: { opacity: 0, x: dx, y: dy },
      animate: { opacity: 1, x: '0%', y: '0%' },
      exit: { opacity: 0, x: dir === 'l' ? '-12%' : dir === 'r' ? '12%' : '0%' },
      transition: { duration: dur, ease: [0.22, 1, 0.36, 1] as const },
    };
  }
  if (transition.type === 'wipe') {
    const initialClip =
      dir === 'l'
        ? 'inset(0 100% 0 0)'
        : dir === 'r'
        ? 'inset(0 0 0 100%)'
        : dir === 'u'
        ? 'inset(100% 0 0 0)'
        : 'inset(0 0 100% 0)';
    return {
      initial: { clipPath: initialClip, opacity: 0.9 },
      animate: { clipPath: 'inset(0 0% 0 0%)', opacity: 1 },
      exit: { opacity: 0 },
      transition: { duration: dur, ease: 'easeOut' as const },
    };
  }
  if (transition.type === 'zoom') {
    return {
      initial: { opacity: 0, scale: 0.88 },
      animate: { opacity: 1, scale: 1 },
      exit: { opacity: 0, scale: 1.05 },
      transition: { duration: dur, ease: [0.22, 1, 0.36, 1] as const },
    };
  }
  if (transition.type === 'wheel') {
    return {
      initial: { opacity: 0, scale: 0.85, rotate: -8 },
      animate: { opacity: 1, scale: 1, rotate: 0 },
      exit: { opacity: 0, scale: 0.95 },
      transition: { duration: dur, ease: [0.22, 1, 0.36, 1] as const },
    };
  }
  if (transition.type === 'flip') {
    return {
      initial: { opacity: 0, rotateY: dir === 'r' ? -75 : 75, scale: 0.92 },
      animate: { opacity: 1, rotateY: 0, scale: 1 },
      exit: { opacity: 0, rotateY: dir === 'r' ? 45 : -45 },
      transition: { duration: dur, ease: [0.22, 1, 0.36, 1] as const },
    };
  }
  if (transition.type === 'split') {
    return {
      initial: { clipPath: 'inset(0 50% 0 50%)', opacity: 0.8 },
      animate: { clipPath: 'inset(0 0% 0 0%)', opacity: 1 },
      exit: { opacity: 0 },
      transition: { duration: dur, ease: 'easeOut' as const },
    };
  }
  return {
    initial: { opacity: 0 },
    animate: { opacity: 1 },
    exit: { opacity: 0 },
    transition: { duration: dur },
  };
};

// Element entrance/emphasis/exit/motionPath variants
const getElementAnimMotion = (
  entr: PptxAnimStep | undefined,
  emph: PptxAnimStep | undefined,
  exit: PptxAnimStep | undefined,
  pathAnim: PptxAnimStep | undefined,
  animStep: number,
  rotDeg = 0,
  baseW = 960,
  baseH = 540
) => {
  const hiddenByEntrance = entr && entr.clickGroup > animStep;
  const hiddenByExit = exit && exit.clickGroup <= animStep;

  const entrStart = entr
    ? (() => {
        const dir = entr.dir || 'd';
        const x = entr.effect === 'fly' ? (dir === 'l' ? -56 : dir === 'r' ? 56 : 0) : 0;
        const y =
          entr.effect === 'fly'
            ? dir === 'u'
              ? -56
              : dir === 'd'
              ? 56
              : 0
            : entr.effect === 'float'
            ? 24
            : 0;
        const scale = entr.effect === 'zoom' || entr.effect === 'bounce' ? 0.65 : 1;
        const clipPath =
          entr.effect === 'wipe'
            ? dir === 'l'
              ? 'inset(0 0 0 100%)'
              : dir === 'u'
              ? 'inset(0 0 100% 0)'
              : dir === 'd'
              ? 'inset(100% 0 0 0)'
              : 'inset(0 100% 0 0)'
            : 'inset(0 0 0 0)';
        return { opacity: 0, x, y, scale, rotate: rotDeg, clipPath };
      })()
    : { opacity: 1, x: 0, y: 0, scale: 1, rotate: rotDeg };

  if (hiddenByEntrance) {
    return {
      initial: entrStart,
      animate: entrStart,
      transition: { duration: 0 },
      pointerEvents: 'none' as const,
    };
  }

  if (hiddenByExit) {
    return {
      initial: { opacity: 1, rotate: rotDeg },
      animate: { opacity: 0, scale: 0.94, rotate: rotDeg },
      transition: { duration: Math.max(0.15, (exit.durationMs || 450) / 1000) },
      pointerEvents: 'none' as const,
    };
  }

  const pathActive = pathAnim && pathAnim.clickGroup <= animStep;
  const pathOffsetX = pathActive && pathAnim.motionDxPct ? (pathAnim.motionDxPct / 100) * baseW : 0;
  const pathOffsetY = pathActive && pathAnim.motionDyPct ? (pathAnim.motionDyPct / 100) * baseH : 0;

  const emphActive = emph && emph.clickGroup === animStep;
  if (emphActive) {
    if (emph.effect === 'spin') {
      return {
        initial: { opacity: 1, x: pathOffsetX, y: pathOffsetY, rotate: rotDeg },
        animate: { opacity: 1, x: pathOffsetX, y: pathOffsetY, rotate: rotDeg + 360 },
        transition: { duration: Math.max(0.4, emph.durationMs / 1000) },
        pointerEvents: 'auto' as const,
      };
    }
    return {
      initial: { opacity: 1, x: pathOffsetX, y: pathOffsetY, scale: 1, rotate: rotDeg },
      animate: { opacity: 1, x: pathOffsetX, y: pathOffsetY, scale: [1, 1.09, 1], rotate: rotDeg },
      transition: { duration: Math.max(0.35, emph.durationMs / 1000) },
      pointerEvents: 'auto' as const,
    };
  }

  if (entr || pathActive) {
    const activeAnim = pathActive && (!entr || pathAnim.clickGroup >= entr.clickGroup) ? pathAnim : entr!;
    const dur = Math.max(0.12, Math.min(2.5, (activeAnim.durationMs || 500) / 1000));
    const delay = activeAnim.trigger === 'onClick' ? 0 : Math.min(2.5, (activeAnim.delayMs || 0) / 1000);
    return {
      initial: entrStart,
      animate: {
        opacity: 1,
        x: pathOffsetX,
        y: pathOffsetY,
        scale: 1,
        rotate: rotDeg,
        clipPath: 'inset(0 0 0 0)',
      },
      transition:
        entr?.effect === 'bounce'
          ? { type: 'spring' as const, stiffness: 320, damping: 14, delay }
          : { duration: dur, delay, ease: [0.22, 1, 0.36, 1] as const },
      pointerEvents: 'auto' as const,
    };
  }

  return {
    initial: { opacity: 1, x: 0, y: 0, rotate: rotDeg },
    animate: { opacity: 1, x: 0, y: 0, rotate: rotDeg },
    transition: { duration: 0.2 },
    pointerEvents: 'auto' as const,
  };
};

interface PptxSlideViewProps {
  pptx: PptxSlideData;
  slideKey: string | number;
  resolveMediaUrl: (id: string) => string;
  animStep: number;
  onSetAnimStep?: (step: number) => void;
  showControls?: boolean;
  maxHeightClass?: string;
  isFullscreen?: boolean;
  fitViewport?: boolean;
}

export const getPptxMaxStep = (pptx?: PptxSlideData) => {
  if (!pptx?.animations?.length) return 0;
  return pptx.animations.reduce((m, a) => Math.max(m, a.clickGroup), 0);
};

// Automatically shrinks font size inside a shape/textbox if browser font metrics exceed the box height
const AutoFitSlideText: React.FC<{
  paragraphs: PptxParagraph[];
  boxHeightPx: number;
  baseFontScale?: number;
}> = ({ paragraphs, boxHeightPx, baseFontScale = 1 }) => {
  const textRef = useRef<HTMLDivElement | null>(null);
  const [fitRatio, setFitRatio] = useState(1);

  useEffect(() => {
    setFitRatio(1);
  }, [paragraphs, boxHeightPx, baseFontScale]);

  useEffect(() => {
    const node = textRef.current;
    if (!node || boxHeightPx <= 0) return;
    if (fitRatio < 1) return;
    const sh = node.scrollHeight;
    if (sh > boxHeightPx + 4) {
      const ratio = Math.max(0.5, Math.min(1, (boxHeightPx / sh) * 0.95));
      if (ratio < 0.98) {
        setFitRatio(ratio);
      }
    }
  }, [paragraphs, boxHeightPx, fitRatio]);

  const effectiveScale = baseFontScale * fitRatio;

  return (
    <div ref={textRef} className="w-full space-y-1">
      {paragraphs.map((p, pIdx) => (
        <div
          key={pIdx}
          style={{
            textAlign: p.align || 'left',
            lineHeight: effectiveScale < 0.85 ? 1.14 : 1.22,
            wordBreak: 'break-word',
            overflowWrap: 'anywhere',
          }}
        >
          {p.bullet && (
            <span
              style={{
                fontSize: `${Math.max(8, (p.runs[0]?.fontSizePt || 18) * 1.12 * effectiveScale)}px`,
                color: p.runs[0]?.color || '#0f172a',
                marginRight: '6px',
              }}
            >
              •
            </span>
          )}
          {p.runs.map((run, rIdx) => {
            if (run.text === '\n') return <br key={rIdx} />;
            const px = Math.max(8, (run.fontSizePt || 18) * 1.16 * effectiveScale);
            return (
              <span
                key={rIdx}
                style={{
                  fontSize: `${px}px`,
                  fontWeight: run.bold ? 700 : 400,
                  fontStyle: run.italic ? 'italic' : 'normal',
                  textDecoration: run.underline ? 'underline' : 'none',
                  color: run.color || '#0f172a',
                  whiteSpace: 'pre-wrap',
                }}
              >
                {run.text}
              </span>
            );
          })}
        </div>
      ))}
    </div>
  );
};

export const PptxSlideView: React.FC<PptxSlideViewProps> = ({
  pptx,
  slideKey,
  resolveMediaUrl,
  animStep,
  onSetAnimStep,
  showControls = false,
  isFullscreen = false,
  fitViewport = true,
}) => {
  const outerRef = useRef<HTMLDivElement | null>(null);
  const baseW = pptx.widthPx || 960;
  const baseH = pptx.heightPx || 540;
  const [availSize, setAvailSize] = useState<{ w: number; h: number }>({ w: baseW, h: baseH });
  const [replayNonce, setReplayNonce] = useState(0);

  const [brokenImgIds, setBrokenImgIds] = useState<Record<string, boolean>>({});

  const maxStep = getPptxMaxStep(pptx);
  const hasAudio = Boolean(pptx.audioIds && pptx.audioIds.length > 0);
  const hasBottomBar = (showControls && Boolean(pptx.transition || maxStep > 0)) || hasAudio;

  // Filter out legacy redundant unplaced-img overlays if the slide already has normal parsed elements
  const hasRealElements = pptx.elements.some(e => !e.id.startsWith('unplaced-img-'));
  const visibleElements = hasRealElements
    ? pptx.elements.filter(e => !e.id.startsWith('unplaced-img-'))
    : pptx.elements;

  useEffect(() => {
    const el = outerRef.current;
    if (!el) return;
    const update = () => {
      const w = el.clientWidth || baseW;
      const vh = typeof window !== 'undefined' ? window.innerHeight : 900;
      const reservedPx = isFullscreen
        ? (hasBottomBar ? 195 : 145)
        : fitViewport
        ? (hasBottomBar ? 255 : 210)
        : 280;
      const maxH = Math.max(240, vh - reservedPx);
      setAvailSize({ w, h: maxH });
    };
    update();
    let ro: ResizeObserver | null = null;
    if (typeof ResizeObserver !== 'undefined') {
      ro = new ResizeObserver(update);
      ro.observe(el);
    }
    window.addEventListener('resize', update);
    return () => {
      ro?.disconnect();
      window.removeEventListener('resize', update);
    };
  }, [baseW, baseH, isFullscreen, fitViewport, hasBottomBar]);

  const uniformScale = Math.max(0.2, Math.min(availSize.w / baseW, availSize.h / baseH));
  const fittedW = Math.round(baseW * uniformScale);
  const fittedH = Math.round(baseH * uniformScale);

  const bgImgUrl = pptx.bgImgId && !brokenImgIds[pptx.bgImgId] ? resolveMediaUrl(pptx.bgImgId) : '';
  const transProps = getSlideTransitionProps(pptx.transition);

  const handleReplay = () => {
    setReplayNonce(n => n + 1);
    onSetAnimStep?.(0);
  };

  return (
    <div ref={outerRef} className="w-full space-y-2.5 flex flex-col items-center">
      <div
        onClick={() => {
          if (onSetAnimStep && maxStep > 0 && animStep < maxStep) {
            onSetAnimStep(animStep + 1);
          }
        }}
        style={{
          width: `${fittedW}px`,
          height: `${fittedH}px`,
          maxWidth: '100%',
        }}
        className={`relative mx-auto overflow-hidden rounded-2xl border border-stone-200 dark:border-slate-800 shadow-md bg-white select-none shrink-0 ${
          onSetAnimStep && animStep < maxStep ? 'cursor-pointer' : ''
        }`}
      >
        <div
          style={{
            width: `${baseW}px`,
            height: `${baseH}px`,
            transform: `scale(${uniformScale})`,
            transformOrigin: 'top left',
          }}
          className="relative overflow-hidden"
        >
          <AnimatePresence mode="wait">
            <motion.div
              key={`${slideKey}-${replayNonce}`}
              initial={transProps.initial}
              animate={transProps.animate}
              exit={transProps.exit}
              transition={transProps.transition}
              style={{
                background: pptx.background || '#ffffff',
                backgroundImage: bgImgUrl ? `url(${bgImgUrl})` : pptx.background?.includes('gradient') ? pptx.background : undefined,
                backgroundSize: '100% 100%',
                backgroundPosition: 'center',
                backgroundRepeat: 'no-repeat',
              }}
              className="absolute inset-0 w-full h-full overflow-hidden"
            >
              {visibleElements.map((el, idx) => {
                const entr = pptx.animations.find(a => a.spid === el.id && a.presetClass === 'entr');
                const emph = pptx.animations.find(a => a.spid === el.id && a.presetClass === 'emph');
                const exit = pptx.animations.find(a => a.spid === el.id && a.presetClass === 'exit');
                const pathAnim = pptx.animations.find(a => a.spid === el.id && a.presetClass === 'path');
                const m = getElementAnimMotion(entr, emph, exit, pathAnim, animStep, el.rotDeg || 0, baseW, baseH);

                const borderRadius =
                  el.geom === 'ellipse'
                    ? '9999px'
                    : el.geom === 'roundRect'
                    ? '12px'
                    : undefined;

                const boxHeightPx = (el.hPct / 100) * baseH;
                const isTransparentTextBox = el.kind === 'shape' && !el.fill && !el.stroke && !el.imgId && !el.vidId;
                const vidUrl = el.vidId ? resolveMediaUrl(el.vidId) : '';
                const imgBroken = Boolean(el.imgId && brokenImgIds[el.imgId]);
                const imgUrl = el.imgId && !imgBroken ? resolveMediaUrl(el.imgId) : '';
                const audUrl = el.audioId ? resolveMediaUrl(el.audioId) : '';

                // If this is a pure picture element whose image failed to decode, skip rendering an empty box
                if (el.kind === 'pic' && !vidUrl && !audUrl && imgBroken && (!el.paragraphs || el.paragraphs.length === 0)) {
                  return null;
                }

                return (
                  <motion.div
                    key={`${el.id}-${idx}-${replayNonce}`}
                    initial={m.initial}
                    animate={m.animate}
                    transition={m.transition}
                    style={{
                      position: 'absolute',
                      left: `${el.xPct}%`,
                      top: `${el.yPct}%`,
                      width: `${el.wPct}%`,
                      height: `${el.hPct}%`,
                      background: el.fill,
                      border: el.stroke ? `${Math.max(1, el.strokeWidthPx || 1)}px solid ${el.stroke}` : undefined,
                      borderRadius,
                      pointerEvents: m.pointerEvents,
                      display: 'flex',
                      flexDirection: 'column',
                      justifyContent: el.valign === 'middle' ? 'center' : el.valign === 'bottom' ? 'flex-end' : 'flex-start',
                      padding: el.kind === 'shape' ? '4px 8px' : undefined,
                      overflow: isTransparentTextBox ? 'visible' : 'hidden',
                    }}
                  >
                    {/* Embedded Video (with optional poster image from PowerPoint) */}
                    {vidUrl ? (
                      <video
                        src={vidUrl}
                        poster={imgUrl || undefined}
                        controls
                        playsInline
                        className="w-full h-full object-contain bg-black/5"
                        onClick={e => e.stopPropagation()}
                      />
                    ) : imgUrl ? (
                      <img
                        src={imgUrl}
                        alt=""
                        onError={() => {
                          if (el.imgId) {
                            setBrokenImgIds(prev => ({ ...prev, [el.imgId!]: true }));
                          }
                        }}
                        style={{
                          transform: `${el.flipH ? 'scaleX(-1)' : ''} ${el.flipV ? 'scaleY(-1)' : ''}`.trim() || undefined,
                        }}
                        className={`${
                          el.paragraphs && el.paragraphs.length > 0
                            ? 'absolute inset-0 w-full h-full object-cover pointer-events-none'
                            : 'w-full h-full object-contain pointer-events-none'
                        }`}
                      />
                    ) : (el.vidId || (el.imgId && !imgBroken)) ? (
                      <div className="w-full h-full flex items-center justify-center text-xs text-stone-400 italic">
                        Đang tải phương tiện...
                      </div>
                    ) : null}

                    {/* Embedded Audio on element */}
                    {audUrl && !vidUrl && (
                      <div className="z-10 p-1" onClick={e => e.stopPropagation()}>
                        <audio src={audUrl} controls className="w-full max-w-[260px] h-8" />
                      </div>
                    )}

                    {/* Table rendering */}
                    {el.kind === 'table' && el.tableRows && (
                      <table className="w-full h-full border-collapse table-fixed text-left">
                        <tbody>
                          {el.tableRows.map((row, rIdx) => (
                            <tr key={rIdx}>
                              {row.map((cell, cIdx) => (
                                <td
                                  key={cIdx}
                                  colSpan={cell.colSpan}
                                  rowSpan={cell.rowSpan}
                                  style={{ background: cell.fill || (rIdx === 0 ? 'rgba(124,58,237,0.12)' : 'rgba(255,255,255,0.85)') }}
                                  className="border border-stone-300 px-2 py-1 align-middle overflow-hidden"
                                >
                                  <AutoFitSlideText
                                    paragraphs={cell.paragraphs}
                                    boxHeightPx={Math.max(24, boxHeightPx / Math.max(1, el.tableRows!.length))}
                                    baseFontScale={0.9}
                                  />
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}

                    {/* Shape / TextBox paragraphs */}
                    {el.paragraphs && el.paragraphs.length > 0 && (
                      <div className="relative z-10 w-full">
                        <AutoFitSlideText
                          paragraphs={el.paragraphs}
                          boxHeightPx={boxHeightPx}
                          baseFontScale={el.fontScale || 1}
                        />
                      </div>
                    )}
                  </motion.div>
                );
              })}
            </motion.div>
          </AnimatePresence>
        </div>
      </div>

      {/* PPTX Transition, Audio & Animation Step Controls */}
      {hasBottomBar && (
        <div
          style={{ width: `${Math.max(300, fittedW)}px`, maxWidth: '100%' }}
          className="flex items-center justify-between flex-wrap gap-2 px-3 py-1.5 rounded-xl bg-stone-50 dark:bg-slate-950 border border-stone-200 dark:border-slate-800 text-xs"
        >
          <div className="flex items-center gap-2 flex-wrap">
            {pptx.transition && (
              <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-purple-100 dark:bg-purple-950 text-purple-700 dark:text-purple-300 font-bold">
                <Layers className="w-3.5 h-3.5" /> {pptx.transition.label || pptx.transition.type}
              </span>
            )}
            {maxStep > 0 && (
              <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-amber-100 dark:bg-amber-950 text-amber-700 font-bold">
                <Sparkles className="w-3.5 h-3.5" /> Hiệu ứng: {animStep} / {maxStep}
              </span>
            )}
            {pptx.audioIds?.map((aid, aIdx) => {
              const u = resolveMediaUrl(aid);
              return u ? (
                <div key={aIdx} className="inline-flex items-center gap-1.5" onClick={e => e.stopPropagation()}>
                  <Volume2 className="w-3.5 h-3.5 text-purple-600" />
                  <audio src={u} controls className="h-7 max-w-[200px]" />
                </div>
              ) : null;
            })}
          </div>

          {showControls && (pptx.transition || maxStep > 0) && (
            <div className="flex items-center gap-1.5 flex-wrap">
              {maxStep > 0 && (
                <>
                  <button
                    type="button"
                    onClick={() => onSetAnimStep?.(Math.max(0, animStep - 1))}
                    disabled={animStep <= 0}
                    className="px-2.5 py-1 rounded-lg bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-700 font-semibold disabled:opacity-40 inline-flex items-center gap-1"
                  >
                    <SkipBack className="w-3.5 h-3.5" /> Bước trước
                  </button>
                  <button
                    type="button"
                    onClick={() => onSetAnimStep?.(Math.min(maxStep, animStep + 1))}
                    disabled={animStep >= maxStep}
                    className="px-2.5 py-1 rounded-lg bg-purple-600 text-white font-bold disabled:opacity-40 inline-flex items-center gap-1"
                  >
                    Bước tiếp <SkipForward className="w-3.5 h-3.5" />
                  </button>
                  <button
                    type="button"
                    onClick={() => onSetAnimStep?.(maxStep)}
                    disabled={animStep >= maxStep}
                    className="px-2.5 py-1 rounded-lg bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-700 font-semibold disabled:opacity-40"
                  >
                    Hiện tất cả
                  </button>
                </>
              )}
              <button
                type="button"
                onClick={handleReplay}
                className="px-2.5 py-1 rounded-lg bg-white dark:bg-slate-900 border border-stone-200 dark:border-slate-700 font-semibold inline-flex items-center gap-1"
              >
                <RotateCcw className="w-3.5 h-3.5" /> Phát lại hiệu ứng
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
};
