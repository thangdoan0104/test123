"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { fabric } from "fabric";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { jsPDF } from "jspdf";
import { saveAs } from "file-saver";

type FieldType = "name" | "alias" | "title" | "custom";
type TextAlign = "left" | "center" | "right" | "justify";
type FillMode = "solid" | "gradient";
type AutoLayoutDirection = "horizontal" | "vertical";
type AutoLayoutAlign = "left" | "center" | "right";
type AutoSpacingMode = "manual" | "fitWidth";
type ExportScope = "active" | "all";

type TemplateState = {
  id: string;
  name: string;
  imageDataUrl: string;
  width: number;
  height: number;
  json?: Record<string, unknown>;
};

type FontAsset = {
  id: string;
  family: string;
  fileName: string;
  dataUrl: string;
};

type GuestRow = Record<string, string | number | boolean | null | undefined>;

type TextObjectSummary = {
  id: string;
  text: string;
  fieldType: FieldType;
  excelColumn: string;
  autoLayoutId: string;
};

type SelectedTextState = {
  id: string;
  text: string;
  fieldType: FieldType;
  excelColumn: string;
  fontFamily: string;
  fontSize: number;
  width: number;
  fill: string;
  fillMode: FillMode;
  gradientFrom: string;
  gradientTo: string;
  gradientAngle: number;
  strokeEnabled: boolean;
  stroke: string;
  strokeWidth: number;
  shadowEnabled: boolean;
  shadowColor: string;
  shadowBlur: number;
  shadowOffsetX: number;
  shadowOffsetY: number;
  textAlign: TextAlign;
  left: number;
  top: number;
  angle: number;
  opacity: number;
  charSpacing: number;
  lineHeight: number;
  autoSpacingMode: AutoSpacingMode;
  minCharSpacing: number;
  maxCharSpacing: number;
  autoLayoutId: string;
  autoLayoutOrder: number;
  autoLayoutDirection: AutoLayoutDirection;
  autoLayoutGap: number;
  autoLayoutAlign: AutoLayoutAlign;
  autoLayoutAnchorX: number;
  autoLayoutAnchorY: number;
};

type TypeStyleConfig = {
  fillMode: FillMode;
  fill: string;
  gradientFrom: string;
  gradientTo: string;
  gradientAngle: number;
  strokeEnabled: boolean;
  stroke: string;
  strokeWidth: number;
  shadowEnabled: boolean;
  shadowColor: string;
  shadowBlur: number;
  shadowOffsetX: number;
  shadowOffsetY: number;
};

type TypeStylesMap = Record<FieldType, TypeStyleConfig>;

type PreviewItem = {
  id: string;
  fileName: string;
  dataUrl: string;
  rowIndex: number;
  templateName: string;
};

type RenderedOutput = {
  fileName: string;
  dataUrl: string;
  width: number;
  height: number;
};

type ProjectFile = {
  version: number;
  savedAt: string;
  templates: TemplateState[];
  fonts: FontAsset[];
  rows: GuestRow[];
  excelColumns: string[];
  typeStyles?: Partial<TypeStylesMap>;
};

const CUSTOM_PROPS = [
  "id",
  "fieldType",
  "excelColumn",
  "fillMode",
  "gradientFrom",
  "gradientTo",
  "gradientAngle",
  "strokeEnabled",
  "strokeColor",
  "strokeWidth",
  "shadowEnabled",
  "shadowColor",
  "shadowBlur",
  "shadowOffsetX",
  "shadowOffsetY",
  "autoSpacingMode",
  "minCharSpacing",
  "maxCharSpacing",
  "autoLayoutId",
  "autoLayoutOrder",
  "autoLayoutDirection",
  "autoLayoutGap",
  "autoLayoutAlign",
  "autoLayoutAnchorX",
  "autoLayoutAnchorY"
];
const STORAGE_KEY = "invitation-card-editor-mvp-project";
const DEFAULT_CANVAS_WIDTH = 900;
const DEFAULT_CANVAS_HEIGHT = 600;

const FIELD_LABELS: Record<FieldType, string> = {
  name: "Ten",
  alias: "Bi danh",
  title: "Chuc vu",
  custom: "Tuy chinh"
};

const DEFAULT_TYPE_STYLES: TypeStylesMap = {
  name: {
    fillMode: "gradient",
    fill: "#ffffff",
    gradientFrom: "#ffffff",
    gradientTo: "#60a5fa",
    gradientAngle: 0,
    strokeEnabled: false,
    stroke: "#0f172a",
    strokeWidth: 0,
    shadowEnabled: false,
    shadowColor: "#0f172a",
    shadowBlur: 12,
    shadowOffsetX: 0,
    shadowOffsetY: 4
  },
  alias: {
    fillMode: "solid",
    fill: "#f8fafc",
    gradientFrom: "#f8fafc",
    gradientTo: "#cbd5e1",
    gradientAngle: 0,
    strokeEnabled: false,
    stroke: "#0f172a",
    strokeWidth: 0,
    shadowEnabled: false,
    shadowColor: "#0f172a",
    shadowBlur: 10,
    shadowOffsetX: 0,
    shadowOffsetY: 3
  },
  title: {
    fillMode: "solid",
    fill: "#ffffff",
    gradientFrom: "#ffffff",
    gradientTo: "#93c5fd",
    gradientAngle: 0,
    strokeEnabled: false,
    stroke: "#0f172a",
    strokeWidth: 0,
    shadowEnabled: false,
    shadowColor: "#0f172a",
    shadowBlur: 10,
    shadowOffsetX: 0,
    shadowOffsetY: 3
  },
  custom: {
    fillMode: "solid",
    fill: "#111111",
    gradientFrom: "#ffffff",
    gradientTo: "#2563eb",
    gradientAngle: 0,
    strokeEnabled: false,
    stroke: "#0f172a",
    strokeWidth: 0,
    shadowEnabled: false,
    shadowColor: "#0f172a",
    shadowBlur: 10,
    shadowOffsetX: 0,
    shadowOffsetY: 3
  }
};


type TextObjectWithMeta = fabric.Textbox & {
  id?: string;
  fieldType?: FieldType;
  excelColumn?: string;
  fillMode?: FillMode;
  gradientFrom?: string;
  gradientTo?: string;
  gradientAngle?: number;
  strokeEnabled?: boolean;
  strokeColor?: string;
  strokeWidth?: number;
  shadowEnabled?: boolean;
  shadowColor?: string;
  shadowBlur?: number;
  shadowOffsetX?: number;
  shadowOffsetY?: number;
  autoSpacingMode?: AutoSpacingMode;
  minCharSpacing?: number;
  maxCharSpacing?: number;
  autoLayoutId?: string;
  autoLayoutOrder?: number;
  autoLayoutDirection?: AutoLayoutDirection;
  autoLayoutGap?: number;
  autoLayoutAlign?: AutoLayoutAlign;
  autoLayoutAnchorX?: number;
  autoLayoutAnchorY?: number;
};

function uid(prefix = "id") {
  return `${prefix}_${Math.random().toString(36).slice(2, 10)}_${Date.now().toString(36)}`;
}

function readFileAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function getImageSize(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
    img.onerror = () => reject(new Error("Cannot read image size"));
    img.src = dataUrl;
  });
}

function isFabricText(obj: fabric.Object): obj is TextObjectWithMeta {
  return obj.type === "i-text" || obj.type === "textbox" || obj.type === "text";
}

function normalizeNumber(value: unknown, fallback = 0) {
  const num = Number(value);
  return Number.isFinite(num) ? num : fallback;
}

function round(value: unknown, digits = 1) {
  const num = normalizeNumber(value, 0);
  const factor = 10 ** digits;
  return Math.round(num * factor) / factor;
}

function sanitizeFontFamily(fileName: string) {
  const base = fileName.replace(/\.[^.]+$/, "");
  return `Font_${base.replace(/[^a-zA-Z0-9_-]+/g, "_")}_${Date.now().toString(36)}`;
}

function safeFileName(input: string) {
  const cleaned = input
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9._-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 90);
  return cleaned || "guest";
}

function dataUrlToBlob(dataUrl: string) {
  const [header, data] = dataUrl.split(",");
  const mime = header.match(/data:(.*?);/)?.[1] || "application/octet-stream";
  const binary = atob(data);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return new Blob([bytes], { type: mime });
}

function displayValue(value: unknown) {
  if (value === null || value === undefined) return "";
  return String(value);
}

function pickGuestName(row: GuestRow, rowIndex: number) {
  const candidates = [
    row.Name,
    row.name,
    row.FullName,
    row.fullName,
    row.full_name,
    row.Ten,
    row["T\u00ean"],
    row["Ho va ten"],
    row["H\u1ecd v\u00e0 t\u00ean"]
  ];
  const found = candidates.find((item) => displayValue(item).trim().length > 0);
  return found ? displayValue(found) : `row-${rowIndex + 1}`;
}


function clampNumber(value: number, min: number, max: number) {
  return Math.min(max, Math.max(min, value));
}

function colorInputValue(value: unknown, fallback = "#111111") {
  const raw = displayValue(value).trim();
  return /^#[0-9a-fA-F]{6}$/.test(raw) ? raw : fallback;
}

function mergeTypeStyles(partial?: Partial<TypeStylesMap>) {
  return {
    name: { ...DEFAULT_TYPE_STYLES.name, ...(partial?.name || {}) },
    alias: { ...DEFAULT_TYPE_STYLES.alias, ...(partial?.alias || {}) },
    title: { ...DEFAULT_TYPE_STYLES.title, ...(partial?.title || {}) },
    custom: { ...DEFAULT_TYPE_STYLES.custom, ...(partial?.custom || {}) }
  } as TypeStylesMap;
}

function getTypeStyle(typeStyles: TypeStylesMap, fieldType?: FieldType) {
  return typeStyles[fieldType || "custom"] || DEFAULT_TYPE_STYLES.custom;
}

function refreshTextDimensions(text: TextObjectWithMeta) {
  const anyText = text as unknown as { initDimensions?: () => void; dirty?: boolean };
  anyText.dirty = true;
  anyText.initDimensions?.();
  text.setCoords();
}

function getTextCenter(text: TextObjectWithMeta) {
  const point = text.getCenterPoint();
  return { x: point.x, y: point.y };
}

function getTextNaturalWidth(text: TextObjectWithMeta) {
  refreshTextDimensions(text);
  const anyText = text as unknown as { __lineWidths?: number[] };
  const lineWidths = Array.isArray(anyText.__lineWidths) ? anyText.__lineWidths : [];
  const measured = lineWidths.length ? Math.max(...lineWidths.map((item) => normalizeNumber(item, 0))) : 0;
  return Math.max(1, measured || normalizeNumber(text.width, normalizeNumber(text.fontSize, 48) * 4));
}

function getAutoLayoutItemSize(text: TextObjectWithMeta) {
  const scaleX = normalizeNumber(text.scaleX, 1) || 1;
  const scaleY = normalizeNumber(text.scaleY, 1) || 1;
  const naturalWidth = getTextNaturalWidth(text) * Math.abs(scaleX);
  const visualHeight = Math.max(1, normalizeNumber(text.height, normalizeNumber(text.fontSize, 48)) * Math.abs(scaleY));
  return { width: naturalWidth, height: visualHeight };
}

function getTextSolidFill(text: TextObjectWithMeta) {
  return colorInputValue(typeof text.fill === "string" ? text.fill : text.gradientFrom, "#111111");
}

function createLinearGradient(width: number, height: number, from: string, to: string, angleValue: number) {
  const safeWidth = Math.max(1, width);
  const safeHeight = Math.max(1, height);
  const angle = (normalizeNumber(angleValue, 0) * Math.PI) / 180;
  const length = Math.sqrt(safeWidth * safeWidth + safeHeight * safeHeight);
  const centerX = safeWidth / 2;
  const centerY = safeHeight / 2;
  const dx = (Math.cos(angle) * length) / 2;
  const dy = (Math.sin(angle) * length) / 2;

  return new fabric.Gradient({
    type: "linear",
    gradientUnits: "pixels",
    coords: {
      x1: centerX - dx,
      y1: centerY - dy,
      x2: centerX + dx,
      y2: centerY + dy
    },
    colorStops: [
      { offset: 0, color: colorInputValue(from, "#ffffff") },
      { offset: 1, color: colorInputValue(to, "#2563eb") }
    ]
  } as any);
}

function createTextGradient(text: TextObjectWithMeta) {
  const width = Math.max(1, normalizeNumber(text.width, 240));
  const height = Math.max(1, normalizeNumber(text.height, normalizeNumber(text.fontSize, 48)));
  return createLinearGradient(width, height, colorInputValue(text.gradientFrom, "#ffffff"), colorInputValue(text.gradientTo, "#2563eb"), normalizeNumber(text.gradientAngle, 0));
}

function applyTextFill(text: TextObjectWithMeta) {
  if (text.fillMode === "gradient") {
    text.set("fill", createTextGradient(text));
  } else {
    text.set("fill", getTextSolidFill(text));
  }
}

function applyTextStroke(text: TextObjectWithMeta) {
  const enabled = !!text.strokeEnabled && normalizeNumber(text.strokeWidth, 0) > 0;
  if (!enabled) {
    text.strokeEnabled = false;
    text.set("strokeWidth", 0);
    text.set("stroke", undefined);
    return;
  }

  text.strokeEnabled = true;
  text.set("strokeWidth", Math.max(0, normalizeNumber(text.strokeWidth, 0)));
  text.set("stroke", colorInputValue(text.strokeColor, "#111111"));
  text.set("paintFirst", "stroke");
  text.set("strokeUniform", true as any);
}
function applyTextShadow(text: TextObjectWithMeta) {
  if (!text.shadowEnabled) {
    text.shadowEnabled = false;
    text.set("shadow", undefined);  // Change from null to undefined
    return;
  }

  text.shadowEnabled = true;
  text.set(
    "shadow",
    new fabric.Shadow({
      color: colorInputValue(text.shadowColor, "#111111"),
      blur: Math.max(0, normalizeNumber(text.shadowBlur, 0)),
      offsetX: normalizeNumber(text.shadowOffsetX, 0),
      offsetY: normalizeNumber(text.shadowOffsetY, 0)
    })
  );
}
  text.shadowEnabled = true;
  text.set(
    "shadow",
    new fabric.Shadow({
      color: colorInputValue(text.shadowColor, "#111111"),
      blur: Math.max(0, normalizeNumber(text.shadowBlur, 0)),
      offsetX: normalizeNumber(text.shadowOffsetX, 0),
      offsetY: normalizeNumber(text.shadowOffsetY, 0)
    })
  );
}

function applyTextAppearance(text: TextObjectWithMeta) {
  applyTextFill(text);
  applyTextStroke(text);
  applyTextShadow(text);
  refreshTextDimensions(text);
}

function applyTypeStyleToText(text: TextObjectWithMeta, style: TypeStyleConfig) {
  text.fillMode = style.fillMode;
  text.gradientFrom = colorInputValue(style.gradientFrom, "#ffffff");
  text.gradientTo = colorInputValue(style.gradientTo, "#2563eb");
  text.gradientAngle = normalizeNumber(style.gradientAngle, 0);
  if (style.fillMode === "solid") {
    text.set("fill", colorInputValue(style.fill, "#111111"));
  } else {
    text.set("fill", colorInputValue(style.fill, "#111111"));
  }
  text.strokeEnabled = !!style.strokeEnabled;
  text.strokeColor = colorInputValue(style.stroke, "#111111");
  text.strokeWidth = normalizeNumber(style.strokeWidth, 0);
  text.shadowEnabled = !!style.shadowEnabled;
  text.shadowColor = colorInputValue(style.shadowColor, "#111111");
  text.shadowBlur = normalizeNumber(style.shadowBlur, 0);
  text.shadowOffsetX = normalizeNumber(style.shadowOffsetX, 0);
  text.shadowOffsetY = normalizeNumber(style.shadowOffsetY, 0);
  applyTextAppearance(text);
}

function applyFieldTypeStylesOnCanvas(canvas: fabric.Canvas | fabric.StaticCanvas, typeStyles: TypeStylesMap) {
  canvas.getObjects().forEach((obj) => {
    if (!isFabricText(obj)) return;
    ensureTextObjectDefaults(obj);
    const style = getTypeStyle(typeStyles, obj.fieldType);
    applyTypeStyleToText(obj, style);
    applyAutoSpacing(obj);
  });
}

function applyAutoSpacing(text: TextObjectWithMeta) {
  if (text.autoSpacingMode !== "fitWidth") return;
  const targetWidth = normalizeNumber(text.width, 0);
  const fontSize = Math.max(1, normalizeNumber(text.fontSize, 48));
  const charCount = displayValue(text.text).replace(/\r?\n/g, "").length;
  if (targetWidth <= 0 || charCount <= 1) return;

  text.set("charSpacing", 0);
  refreshTextDimensions(text);
  const naturalWidth = getTextNaturalWidth(text);
  const denominator = fontSize * Math.max(1, charCount - 1);
  const minSpacing = normalizeNumber(text.minCharSpacing, -180);
  const maxSpacing = normalizeNumber(text.maxCharSpacing, 900);
  const computedSpacing = clampNumber(((targetWidth - naturalWidth) / denominator) * 1000, minSpacing, maxSpacing);
  text.set("charSpacing", Math.round(computedSpacing));
  refreshTextDimensions(text);
}

function ensureTextObjectDefaults(text: TextObjectWithMeta) {
  if (!text.id) text.id = uid("text");
  if (!text.fieldType) text.fieldType = "custom";
  if (text.excelColumn === undefined) text.excelColumn = "";
  if (!text.fillMode) text.fillMode = typeof text.fill === "string" ? "solid" : "gradient";
  if (!text.gradientFrom) text.gradientFrom = getTextSolidFill(text);
  if (!text.gradientTo) text.gradientTo = "#2563eb";
  if (text.gradientAngle === undefined) text.gradientAngle = 0;
  if (text.strokeEnabled === undefined) text.strokeEnabled = normalizeNumber(text.strokeWidth, 0) > 0;
  if (!text.strokeColor) text.strokeColor = colorInputValue(text.stroke, "#111111");
  if (text.strokeWidth === undefined) text.strokeWidth = normalizeNumber(text.strokeWidth, 0);
  if (text.shadowEnabled === undefined) text.shadowEnabled = !!text.shadow;
  const currentShadow = text.shadow as fabric.Shadow | null | undefined;
  if (!text.shadowColor) text.shadowColor = colorInputValue(currentShadow?.color, "#111111");
  if (text.shadowBlur === undefined) text.shadowBlur = normalizeNumber(currentShadow?.blur, 0);
  if (text.shadowOffsetX === undefined) text.shadowOffsetX = normalizeNumber(currentShadow?.offsetX, 0);
  if (text.shadowOffsetY === undefined) text.shadowOffsetY = normalizeNumber(currentShadow?.offsetY, 0);
  if (!text.autoSpacingMode) text.autoSpacingMode = "manual";
  if (text.minCharSpacing === undefined) text.minCharSpacing = -180;
  if (text.maxCharSpacing === undefined) text.maxCharSpacing = 900;
  if (!text.autoLayoutDirection) text.autoLayoutDirection = "horizontal";
  if (!text.autoLayoutAlign) text.autoLayoutAlign = "center";
  if (text.autoLayoutGap === undefined) text.autoLayoutGap = 16;
  if (text.autoLayoutOrder === undefined) text.autoLayoutOrder = 1;

  const currentWidth = normalizeNumber(text.width, 0);
  if (currentWidth < 20) {
    const size = normalizeNumber(text.fontSize, 48);
    text.set("width", Math.max(220, size * 6));
  }

  text.set({ editable: true, objectCaching: false });
  refreshTextDimensions(text);
}

function cloneTextObjectAsTextbox(obj: fabric.Object) {
  const source = obj as TextObjectWithMeta;
  const fontSize = normalizeNumber(source.fontSize, 48);
  const sourceWidth = normalizeNumber(source.width, 0);
  const textbox = new fabric.Textbox(displayValue(source.text), {
    left: source.left,
    top: source.top,
    width: Math.max(220, sourceWidth * 1.35, fontSize * 6),
    originX: (source.originX || "center") as any,
    originY: (source.originY || "center") as any,
    angle: source.angle,
    scaleX: source.scaleX,
    scaleY: source.scaleY,
    opacity: source.opacity,
    fontFamily: source.fontFamily,
    fontSize: source.fontSize,
    fontWeight: source.fontWeight as any,
    fontStyle: source.fontStyle,
    underline: source.underline,
    linethrough: source.linethrough,
    overline: source.overline,
    fill: source.fill as any,
    stroke: source.stroke as any,
    strokeWidth: source.strokeWidth,
    shadow: source.shadow as any,
    textAlign: (source.textAlign || "center") as TextAlign,
    charSpacing: source.charSpacing,
    lineHeight: source.lineHeight,
    editable: true,
    objectCaching: false
  }) as TextObjectWithMeta;

  CUSTOM_PROPS.forEach((prop) => {
    (textbox as unknown as Record<string, unknown>)[prop] = (source as unknown as Record<string, unknown>)[prop];
  });
  ensureTextObjectDefaults(textbox);
  return textbox;
}

function normalizeCanvasTextObjects(canvas: fabric.Canvas | fabric.StaticCanvas) {
  const objects = [...canvas.getObjects()];
  objects.forEach((obj, index) => {
    if (!isFabricText(obj)) return;
    let text = obj as TextObjectWithMeta;
    if (obj.type !== "textbox") {
      const textbox = cloneTextObjectAsTextbox(obj);
      canvas.remove(obj);
      canvas.insertAt(textbox, index, false);
      text = textbox;
    }
    ensureTextObjectDefaults(text);
    applyAutoSpacing(text);
    applyTextAppearance(text);
  });
}

function groupTextObjectsByAutoLayout(canvas: fabric.Canvas | fabric.StaticCanvas) {
  const groups = new Map<string, TextObjectWithMeta[]>();
  canvas.getObjects().forEach((obj) => {
    if (!isFabricText(obj)) return;
    const groupId = displayValue(obj.autoLayoutId).trim();
    if (!groupId) return;
    if (!groups.has(groupId)) groups.set(groupId, []);
    groups.get(groupId)?.push(obj);
  });
  return groups;
}

function updateAutoLayoutAnchorFromBounds(canvas: fabric.Canvas | fabric.StaticCanvas, groupId: string) {
  const items = groupTextObjectsByAutoLayout(canvas).get(groupId);
  if (!items?.length) return;
  const leader = [...items].sort((a, b) => normalizeNumber(a.autoLayoutOrder, 1) - normalizeNumber(b.autoLayoutOrder, 1))[0];
  const align = leader.autoLayoutAlign || "center";
  const rects = items.map((item) => item.getBoundingRect(true, true));
  const left = Math.min(...rects.map((rect) => rect.left));
  const top = Math.min(...rects.map((rect) => rect.top));
  const right = Math.max(...rects.map((rect) => rect.left + rect.width));
  const bottom = Math.max(...rects.map((rect) => rect.top + rect.height));
  const anchorX = align === "left" ? left : align === "right" ? right : (left + right) / 2;
  const anchorY = (top + bottom) / 2;
  items.forEach((item) => {
    item.autoLayoutAnchorX = round(anchorX, 2);
    item.autoLayoutAnchorY = round(anchorY, 2);
  });
}

function applyAutoLayouts(canvas: fabric.Canvas | fabric.StaticCanvas) {
  const groups = groupTextObjectsByAutoLayout(canvas);
  groups.forEach((items) => {
    const ordered = [...items].sort((a, b) => normalizeNumber(a.autoLayoutOrder, 1) - normalizeNumber(b.autoLayoutOrder, 1));
    const leader = ordered[0];
    const direction = leader.autoLayoutDirection || "horizontal";
    const gap = normalizeNumber(leader.autoLayoutGap, 16);
    const align = leader.autoLayoutAlign || "center";
    const fallbackCenter = getTextCenter(leader);
    const anchorX = normalizeNumber(leader.autoLayoutAnchorX, fallbackCenter.x);
    const anchorY = normalizeNumber(leader.autoLayoutAnchorY, fallbackCenter.y);

    ordered.forEach((item, index) => {
      item.autoLayoutDirection = direction;
      item.autoLayoutGap = gap;
      item.autoLayoutAlign = align;
      item.autoLayoutAnchorX = anchorX;
      item.autoLayoutAnchorY = anchorY;
      if (item.autoLayoutOrder === undefined) item.autoLayoutOrder = index + 1;
      applyAutoSpacing(item);
      applyTextAppearance(item);
    });

    const sizes = ordered.map(getAutoLayoutItemSize);
    if (direction === "horizontal") {
      const totalWidth = sizes.reduce((sum, size) => sum + size.width, 0) + gap * Math.max(0, ordered.length - 1);
      let cursorX = align === "left" ? anchorX : align === "right" ? anchorX - totalWidth : anchorX - totalWidth / 2;
      ordered.forEach((item, index) => {
        const size = sizes[index];
        const centerX = cursorX + size.width / 2;
        item.setPositionByOrigin(new fabric.Point(centerX, anchorY), "center", "center");
        item.setCoords();
        cursorX += size.width + gap;
      });
      return;
    }

    const totalHeight = sizes.reduce((sum, size) => sum + size.height, 0) + gap * Math.max(0, ordered.length - 1);
    let cursorY = anchorY - totalHeight / 2;
    ordered.forEach((item, index) => {
      const size = sizes[index];
      const centerX = align === "left" ? anchorX + size.width / 2 : align === "right" ? anchorX - size.width / 2 : anchorX;
      const centerY = cursorY + size.height / 2;
      item.setPositionByOrigin(new fabric.Point(centerX, centerY), "center", "center");
      item.setCoords();
      cursorY += size.height + gap;
    });
  });
  canvas.requestRenderAll();
}

function refreshCanvasTextLayout(canvas: fabric.Canvas | fabric.StaticCanvas) {
  normalizeCanvasTextObjects(canvas);
  applyAutoLayouts(canvas);
  canvas.requestRenderAll();
}

function setCanvasBackground(canvas: fabric.StaticCanvas | fabric.Canvas, template: TemplateState) {
  return new Promise<void>((resolve, reject) => {
    fabric.Image.fromURL(
      template.imageDataUrl,
      (img) => {
        if (!img) {
          reject(new Error("Cannot load background image"));
          return;
        }
        const imgWidth = img.width || template.width;
        const imgHeight = img.height || template.height;
        img.set({
          left: 0,
          top: 0,
          originX: "left",
          originY: "top",
          selectable: false,
          evented: false
        });
        img.scaleX = template.width / imgWidth;
        img.scaleY = template.height / imgHeight;
        canvas.setBackgroundImage(img, () => {
          canvas.renderAll();
          resolve();
        });
      },
      { crossOrigin: "anonymous" }
    );
  });
}

async function registerFontAsset(font: FontAsset) {
  if (typeof window === "undefined" || !("FontFace" in window)) return;
  const fontFace = new FontFace(font.family, `url(${font.dataUrl})`);
  await fontFace.load();
  document.fonts.add(fontFace);
}

async function registerFontAssets(fonts: FontAsset[]) {
  await Promise.all(fonts.map((font) => registerFontAsset(font).catch(() => undefined)));
}

async function loadTemplateIntoStaticCanvas(template: TemplateState, typeStyles: TypeStylesMap) {
  const element = document.createElement("canvas");
  const canvas = new fabric.StaticCanvas(element, {
    width: template.width,
    height: template.height,
    renderOnAddRemove: false,
    backgroundColor: "#ffffff"
  });

  if (template.json) {
    await new Promise<void>((resolve) => {
      canvas.loadFromJSON(template.json, () => resolve());
    });
    canvas.setWidth(template.width);
    canvas.setHeight(template.height);
    if (!canvas.backgroundImage) {
      await setCanvasBackground(canvas, template);
    }
  } else {
    await setCanvasBackground(canvas, template);
  }

  applyFieldTypeStylesOnCanvas(canvas, typeStyles);
  refreshCanvasTextLayout(canvas);
  return canvas;
}

export default function InvitationDesigner() {
  const canvasElementRef = useRef<HTMLCanvasElement | null>(null);
  const fabricCanvasRef = useRef<fabric.Canvas | null>(null);
  const loadedTemplateIdRef = useRef<string | null>(null);

  const [templates, setTemplates] = useState<TemplateState[]>([]);
  const [activeTemplateId, setActiveTemplateId] = useState<string | null>(null);
  const [fonts, setFonts] = useState<FontAsset[]>([]);
  const [excelColumns, setExcelColumns] = useState<string[]>([]);
  const [rows, setRows] = useState<GuestRow[]>([]);
  const [textObjects, setTextObjects] = useState<TextObjectSummary[]>([]);
  const [selectedText, setSelectedText] = useState<SelectedTextState | null>(null);
  const [previewItems, setPreviewItems] = useState<PreviewItem[]>([]);
  const [busyMessage, setBusyMessage] = useState<string | null>(null);
  const [exportScale, setExportScale] = useState(2);
  const [jpegQuality, setJpegQuality] = useState(0.95);
  const [previewLimit, setPreviewLimit] = useState(30);
  const [previewScope, setPreviewScope] = useState<ExportScope>("active");
  const [zoom, setZoom] = useState(1);
  const [typeStyles, setTypeStyles] = useState<TypeStylesMap>(() => mergeTypeStyles());

  const activeTemplate = useMemo(
    () => templates.find((template) => template.id === activeTemplateId) || null,
    [templates, activeTemplateId]
  );

  const autoLayoutGroups = useMemo(
    () => Array.from(new Set(textObjects.map((item) => item.autoLayoutId).filter(Boolean))),
    [textObjects]
  );

  const refreshTextObjects = useCallback(() => {
    const canvas = fabricCanvasRef.current;
    if (!canvas) {
      setTextObjects([]);
      return;
    }
    const summaries = canvas
      .getObjects()
      .filter(isFabricText)
      .map((obj) => {
        ensureTextObjectDefaults(obj);
        return {
          id: obj.id || "",
          text: displayValue(obj.text),
          fieldType: obj.fieldType || "custom",
          excelColumn: obj.excelColumn || "",
          autoLayoutId: obj.autoLayoutId || ""
        };
      });
    setTextObjects(summaries);
  }, []);

  const syncSelection = useCallback(() => {
    const canvas = fabricCanvasRef.current;
    const active = canvas?.getActiveObject();
    if (!active || !isFabricText(active)) {
      setSelectedText(null);
      return;
    }

    const anyActive = active as TextObjectWithMeta;
    ensureTextObjectDefaults(anyActive);
    const center = getTextCenter(anyActive);

    setSelectedText({
      id: anyActive.id || "",
      text: displayValue(anyActive.text),
      fieldType: anyActive.fieldType || "custom",
      excelColumn: anyActive.excelColumn || "",
      fontFamily: displayValue(anyActive.fontFamily || "Arial"),
      fontSize: normalizeNumber(anyActive.fontSize, 48),
      width: round(anyActive.width),
      fill: getTextSolidFill(anyActive),
      fillMode: anyActive.fillMode || "solid",
      gradientFrom: colorInputValue(anyActive.gradientFrom, getTextSolidFill(anyActive)),
      gradientTo: colorInputValue(anyActive.gradientTo, "#2563eb"),
      gradientAngle: normalizeNumber(anyActive.gradientAngle, 0),
      strokeEnabled: !!anyActive.strokeEnabled && normalizeNumber(anyActive.strokeWidth, 0) > 0,
      stroke: colorInputValue(anyActive.strokeColor, "#111111"),
      strokeWidth: normalizeNumber(anyActive.strokeWidth, 0),
      shadowEnabled: !!anyActive.shadowEnabled,
      shadowColor: colorInputValue(anyActive.shadowColor, "#111111"),
      shadowBlur: normalizeNumber(anyActive.shadowBlur, 0),
      shadowOffsetX: normalizeNumber(anyActive.shadowOffsetX, 0),
      shadowOffsetY: normalizeNumber(anyActive.shadowOffsetY, 0),
      textAlign: (anyActive.textAlign || "center") as TextAlign,
      left: round(anyActive.left),
      top: round(anyActive.top),
      angle: round(anyActive.angle),
      opacity: normalizeNumber(anyActive.opacity, 1),
      charSpacing: normalizeNumber(anyActive.charSpacing, 0),
      lineHeight: normalizeNumber(anyActive.lineHeight, 1.16),
      autoSpacingMode: anyActive.autoSpacingMode || "manual",
      minCharSpacing: normalizeNumber(anyActive.minCharSpacing, -180),
      maxCharSpacing: normalizeNumber(anyActive.maxCharSpacing, 900),
      autoLayoutId: anyActive.autoLayoutId || "",
      autoLayoutOrder: normalizeNumber(anyActive.autoLayoutOrder, 1),
      autoLayoutDirection: anyActive.autoLayoutDirection || "horizontal",
      autoLayoutGap: normalizeNumber(anyActive.autoLayoutGap, 16),
      autoLayoutAlign: anyActive.autoLayoutAlign || "center",
      autoLayoutAnchorX: normalizeNumber(anyActive.autoLayoutAnchorX, center.x),
      autoLayoutAnchorY: normalizeNumber(anyActive.autoLayoutAnchorY, center.y)
    });
  }, []);

  const loadTemplateToCanvas = useCallback(
    async (template: TemplateState) => {
      const canvas = fabricCanvasRef.current;
      if (!canvas) return;
      setBusyMessage("Loading template...");
      canvas.clear();
      canvas.setWidth(template.width);
      canvas.setHeight(template.height);
      canvas.setBackgroundColor("#ffffff", () => undefined);

      if (template.json) {
        await new Promise<void>((resolve) => {
          canvas.loadFromJSON(template.json, () => resolve());
        });
        canvas.setWidth(template.width);
        canvas.setHeight(template.height);
        if (!canvas.backgroundImage) {
          await setCanvasBackground(canvas, template);
        }
      } else {
        await setCanvasBackground(canvas, template);
      }

      applyFieldTypeStylesOnCanvas(canvas, typeStyles);
      refreshCanvasTextLayout(canvas);
      canvas.discardActiveObject();
      canvas.requestRenderAll();
      refreshTextObjects();
      syncSelection();
      setBusyMessage(null);
    },
    [refreshTextObjects, syncSelection, typeStyles]
  );

  const getTemplatesWithCurrentCanvas = useCallback(() => {
    const canvas = fabricCanvasRef.current;
    if (!canvas || !activeTemplateId) return templates;
    const currentJson = canvas.toJSON(CUSTOM_PROPS) as Record<string, unknown>;
    return templates.map((template) =>
      template.id === activeTemplateId
        ? {
            ...template,
            width: canvas.getWidth(),
            height: canvas.getHeight(),
            json: currentJson
          }
        : template
    );
  }, [templates, activeTemplateId]);

  useEffect(() => {
    const element = canvasElementRef.current;
    if (!element || fabricCanvasRef.current) return;

    const canvas = new fabric.Canvas(element, {
      width: DEFAULT_CANVAS_WIDTH,
      height: DEFAULT_CANVAS_HEIGHT,
      preserveObjectStacking: true,
      selection: true,
      backgroundColor: "#ffffff"
    });

    fabric.Object.prototype.transparentCorners = false;
    fabric.Object.prototype.cornerStyle = "circle";
    fabric.Object.prototype.cornerSize = 11;
    fabric.Object.prototype.borderColor = "#2563eb";
    fabric.Object.prototype.cornerColor = "#2563eb";

    canvas.on("selection:created", () => {
      syncSelection();
      refreshTextObjects();
    });
    canvas.on("selection:updated", () => {
      syncSelection();
      refreshTextObjects();
    });
    canvas.on("selection:cleared", () => setSelectedText(null));
    canvas.on("object:modified", (event) => {
      const target = event.target;
      if (target && isFabricText(target)) {
        if (target.autoLayoutId) updateAutoLayoutAnchorFromBounds(canvas, target.autoLayoutId);
        refreshCanvasTextLayout(canvas);
      }
      syncSelection();
      refreshTextObjects();
    });
    canvas.on("object:moving", syncSelection);
    canvas.on("object:scaling", syncSelection);
    canvas.on("object:rotating", syncSelection);
    canvas.on("text:changed", (event) => {
      const target = event.target;
      if (target && isFabricText(target)) {
        applyAutoSpacing(target);
        applyTextAppearance(target);
        if (target.autoLayoutId) applyAutoLayouts(canvas);
      }
      syncSelection();
      refreshTextObjects();
    });

    fabricCanvasRef.current = canvas;
    return () => {
      canvas.dispose();
      fabricCanvasRef.current = null;
    };
  }, [refreshTextObjects, syncSelection]);

  useEffect(() => {
    if (!activeTemplate || !fabricCanvasRef.current) return;
    if (loadedTemplateIdRef.current === activeTemplate.id) return;
    loadedTemplateIdRef.current = activeTemplate.id;
    void loadTemplateToCanvas(activeTemplate);
  }, [activeTemplate, loadTemplateToCanvas]);

  useEffect(() => {
    const canvas = fabricCanvasRef.current;
    if (!canvas) return;
    canvas.setZoom(zoom);
    canvas.requestRenderAll();
  }, [zoom]);

  useEffect(() => {
    const canvas = fabricCanvasRef.current;
    if (!canvas) return;
    applyFieldTypeStylesOnCanvas(canvas, typeStyles);
    refreshCanvasTextLayout(canvas);
    syncSelection();
    refreshTextObjects();
  }, [typeStyles, refreshTextObjects, syncSelection]);

  function switchTemplate(templateId: string) {
    const nextTemplates = getTemplatesWithCurrentCanvas();
    setTemplates(nextTemplates);
    loadedTemplateIdRef.current = null;
    setActiveTemplateId(templateId);
    setSelectedText(null);
    setPreviewItems([]);
  }

  async function handleUploadImages(files: FileList | null) {
    if (!files?.length) return;
    setBusyMessage("Reading images...");
    try {
      const nextTemplates = [] as TemplateState[];
      for (const file of Array.from(files)) {
        const dataUrl = await readFileAsDataUrl(file);
        const size = await getImageSize(dataUrl);
        nextTemplates.push({
          id: uid("template"),
          name: file.name,
          imageDataUrl: dataUrl,
          width: size.width,
          height: size.height
        });
      }

      const current = getTemplatesWithCurrentCanvas();
      const merged = [...current, ...nextTemplates];
      setTemplates(merged);
      if (!activeTemplateId && nextTemplates.length > 0) {
        loadedTemplateIdRef.current = null;
        setActiveTemplateId(nextTemplates[0].id);
      }
    } finally {
      setBusyMessage(null);
    }
  }

  async function handleUploadFont(files: FileList | null) {
    if (!files?.length) return;
    setBusyMessage("Loading fonts...");
    try {
      const newFonts: FontAsset[] = [];
      for (const file of Array.from(files)) {
        const dataUrl = await readFileAsDataUrl(file);
        const font: FontAsset = {
          id: uid("font"),
          family: sanitizeFontFamily(file.name),
          fileName: file.name,
          dataUrl
        };
        await registerFontAsset(font);
        newFonts.push(font);
      }
      setFonts((prev) => [...prev, ...newFonts]);
    } finally {
      setBusyMessage(null);
    }
  }

  async function handleUploadSpreadsheet(file: File | null) {
    if (!file) return;
    setBusyMessage("Reading Excel/CSV...");
    try {
      const isCsv = file.name.toLowerCase().endsWith(".csv");
      const workbook = isCsv
        ? XLSX.read(await file.text(), { type: "string" })
        : XLSX.read(await file.arrayBuffer(), { type: "array" });
      const firstSheetName = workbook.SheetNames[0];
      if (!firstSheetName) throw new Error("Workbook has no sheets");
      const sheet = workbook.Sheets[firstSheetName];
      const data = XLSX.utils.sheet_to_json<GuestRow>(sheet, { defval: "" });
      const columns = data.length ? Object.keys(data[0]) : [];
      setRows(data);
      setExcelColumns(columns);
      setPreviewItems([]);
    } finally {
      setBusyMessage(null);
    }
  }

  function addTextBox() {
    const canvas = fabricCanvasRef.current;
    if (!canvas || !activeTemplate) return;
    const fontSize = Math.max(24, Math.round(activeTemplate.width / 18));
    const text = new fabric.Textbox("Guest Name", {
      left: activeTemplate.width / 2,
      top: activeTemplate.height / 2,
      width: Math.max(260, activeTemplate.width * 0.48),
      originX: "center",
      originY: "center",
      fontFamily: fonts[0]?.family || "Arial",
      fontSize,
      fill: "#111111",
      textAlign: "center",
      editable: true,
      objectCaching: false
    }) as TextObjectWithMeta;
    text.id = uid("text");
    text.fieldType = "name";
    text.excelColumn = excelColumns[0] || "";
    text.fillMode = "solid";
    text.gradientFrom = "#ffffff";
    text.gradientTo = "#2563eb";
    text.gradientAngle = 0;
    text.strokeEnabled = false;
    text.strokeColor = "#111111";
    text.strokeWidth = 0;
    text.shadowEnabled = false;
    text.shadowColor = "#111111";
    text.shadowBlur = 10;
    text.shadowOffsetX = 0;
    text.shadowOffsetY = 3;
    text.autoSpacingMode = "manual";
    text.minCharSpacing = -180;
    text.maxCharSpacing = 900;
    ensureTextObjectDefaults(text);
    applyTypeStyleToText(text, getTypeStyle(typeStyles, text.fieldType));
    canvas.add(text);
    canvas.setActiveObject(text);
    text.enterEditing();
    text.selectAll();
    canvas.requestRenderAll();
    syncSelection();
    refreshTextObjects();
  }

  function selectTextObject(id: string) {
    const canvas = fabricCanvasRef.current;
    if (!canvas) return;
    const target = canvas.getObjects().find((obj) => (obj as { id?: string }).id === id);
    if (!target) return;
    canvas.setActiveObject(target);
    canvas.requestRenderAll();
    syncSelection();
  }

  function updateActiveText(patch: Partial<SelectedTextState>) {
    const canvas = fabricCanvasRef.current;
    const active = canvas?.getActiveObject();
    if (!canvas || !active || !isFabricText(active)) return;
    const text = active as TextObjectWithMeta;
    ensureTextObjectDefaults(text);

    if (patch.text !== undefined) text.set("text", patch.text);
    if (patch.fontFamily !== undefined) text.set("fontFamily", patch.fontFamily);
    if (patch.fontSize !== undefined) text.set("fontSize", normalizeNumber(patch.fontSize, 48));
    if (patch.width !== undefined) text.set("width", Math.max(20, normalizeNumber(patch.width, normalizeNumber(text.width, 240))));
    if (patch.fillMode !== undefined) text.fillMode = patch.fillMode;
    if (patch.fill !== undefined) {
      text.fillMode = "solid";
      text.set("fill", colorInputValue(patch.fill, getTextSolidFill(text)));
    }
    if (patch.gradientFrom !== undefined) text.gradientFrom = colorInputValue(patch.gradientFrom, "#ffffff");
    if (patch.gradientTo !== undefined) text.gradientTo = colorInputValue(patch.gradientTo, "#2563eb");
    if (patch.gradientAngle !== undefined) text.gradientAngle = normalizeNumber(patch.gradientAngle, 0);
    if (patch.strokeEnabled !== undefined) text.strokeEnabled = patch.strokeEnabled;
    if (patch.stroke !== undefined) text.strokeColor = colorInputValue(patch.stroke, "#111111");
    if (patch.strokeWidth !== undefined) text.strokeWidth = normalizeNumber(patch.strokeWidth, 0);
    if (patch.shadowEnabled !== undefined) text.shadowEnabled = patch.shadowEnabled;
    if (patch.shadowColor !== undefined) text.shadowColor = colorInputValue(patch.shadowColor, "#111111");
    if (patch.shadowBlur !== undefined) text.shadowBlur = normalizeNumber(patch.shadowBlur, 0);
    if (patch.shadowOffsetX !== undefined) text.shadowOffsetX = normalizeNumber(patch.shadowOffsetX, 0);
    if (patch.shadowOffsetY !== undefined) text.shadowOffsetY = normalizeNumber(patch.shadowOffsetY, 0);
    if (patch.textAlign !== undefined) text.set("textAlign", patch.textAlign);
    if (patch.left !== undefined) text.set("left", normalizeNumber(patch.left, 0));
    if (patch.top !== undefined) text.set("top", normalizeNumber(patch.top, 0));
    if (patch.angle !== undefined) text.set("angle", normalizeNumber(patch.angle, 0));
    if (patch.opacity !== undefined) text.set("opacity", normalizeNumber(patch.opacity, 1));
    if (patch.charSpacing !== undefined) text.set("charSpacing", normalizeNumber(patch.charSpacing, 0));
    if (patch.lineHeight !== undefined) text.set("lineHeight", normalizeNumber(patch.lineHeight, 1.16));
    if (patch.fieldType !== undefined) {
      text.fieldType = patch.fieldType;
      applyTypeStyleToText(text, getTypeStyle(typeStyles, patch.fieldType));
    }
    if (patch.excelColumn !== undefined) text.excelColumn = patch.excelColumn;
    if (patch.autoSpacingMode !== undefined) text.autoSpacingMode = patch.autoSpacingMode;
    if (patch.minCharSpacing !== undefined) text.minCharSpacing = normalizeNumber(patch.minCharSpacing, -180);
    if (patch.maxCharSpacing !== undefined) text.maxCharSpacing = normalizeNumber(patch.maxCharSpacing, 900);

    if (patch.autoLayoutId !== undefined) {
      const nextGroupId = patch.autoLayoutId.trim();
      const existingGroupObject = canvas
        .getObjects()
        .find((obj) => obj !== text && isFabricText(obj) && obj.autoLayoutId === nextGroupId) as TextObjectWithMeta | undefined;
      text.autoLayoutId = nextGroupId;
      if (nextGroupId && existingGroupObject) {
        text.autoLayoutDirection = existingGroupObject.autoLayoutDirection || "horizontal";
        text.autoLayoutGap = normalizeNumber(existingGroupObject.autoLayoutGap, 16);
        text.autoLayoutAlign = existingGroupObject.autoLayoutAlign || "center";
        text.autoLayoutAnchorX = normalizeNumber(existingGroupObject.autoLayoutAnchorX, getTextCenter(existingGroupObject).x);
        text.autoLayoutAnchorY = normalizeNumber(existingGroupObject.autoLayoutAnchorY, getTextCenter(existingGroupObject).y);
      } else if (nextGroupId) {
        const center = getTextCenter(text);
        text.autoLayoutAnchorX = center.x;
        text.autoLayoutAnchorY = center.y;
      }
    }
    if (patch.autoLayoutOrder !== undefined) text.autoLayoutOrder = normalizeNumber(patch.autoLayoutOrder, 1);
    if (patch.autoLayoutDirection !== undefined) text.autoLayoutDirection = patch.autoLayoutDirection;
    if (patch.autoLayoutGap !== undefined) text.autoLayoutGap = normalizeNumber(patch.autoLayoutGap, 16);
    if (patch.autoLayoutAlign !== undefined) text.autoLayoutAlign = patch.autoLayoutAlign;
    if (patch.autoLayoutAnchorX !== undefined) text.autoLayoutAnchorX = normalizeNumber(patch.autoLayoutAnchorX, getTextCenter(text).x);
    if (patch.autoLayoutAnchorY !== undefined) text.autoLayoutAnchorY = normalizeNumber(patch.autoLayoutAnchorY, getTextCenter(text).y);

    const groupId = displayValue(text.autoLayoutId).trim();
    if (groupId) {
      canvas.getObjects().forEach((obj) => {
        if (!isFabricText(obj) || obj.autoLayoutId !== groupId) return;
        if (patch.autoLayoutDirection !== undefined) obj.autoLayoutDirection = text.autoLayoutDirection;
        if (patch.autoLayoutGap !== undefined) obj.autoLayoutGap = text.autoLayoutGap;
        if (patch.autoLayoutAlign !== undefined) obj.autoLayoutAlign = text.autoLayoutAlign;
        if (patch.autoLayoutAnchorX !== undefined) obj.autoLayoutAnchorX = text.autoLayoutAnchorX;
        if (patch.autoLayoutAnchorY !== undefined) obj.autoLayoutAnchorY = text.autoLayoutAnchorY;
      });
    }

    refreshTextDimensions(text);
    applyAutoSpacing(text);
    applyTextAppearance(text);
    if (groupId) applyAutoLayouts(canvas);
    else canvas.requestRenderAll();
    syncSelection();
    refreshTextObjects();
  }

  function deleteActiveText() {
    const canvas = fabricCanvasRef.current;
    const active = canvas?.getActiveObject();
    if (!canvas || !active || !isFabricText(active)) return;
    canvas.remove(active);
    canvas.discardActiveObject();
    canvas.requestRenderAll();
    setSelectedText(null);
    refreshTextObjects();
  }

  function duplicateActiveText() {
    const canvas = fabricCanvasRef.current;
    const active = canvas?.getActiveObject();
    if (!canvas || !active || !isFabricText(active)) return;
    active.clone((clonedObject: fabric.Object) => {
      const clone = clonedObject as TextObjectWithMeta;
      clone.set({ left: (active.left || 0) + 24, top: (active.top || 0) + 24 });
      clone.id = uid("text");
      clone.fieldType = active.fieldType || "custom";
      clone.excelColumn = active.excelColumn || "";
      clone.autoLayoutOrder = normalizeNumber(active.autoLayoutOrder, 1) + 1;
      ensureTextObjectDefaults(clone);
      applyTextAppearance(clone);
      canvas.add(clone);
      canvas.setActiveObject(clone);
      if (clone.autoLayoutId) applyAutoLayouts(canvas);
      else canvas.requestRenderAll();
      syncSelection();
      refreshTextObjects();
    }, CUSTOM_PROPS);
  }

  function createAutoLayoutGroupForActiveText() {
    const canvas = fabricCanvasRef.current;
    const active = canvas?.getActiveObject();
    if (!canvas || !active || !isFabricText(active)) return;
    const text = active as TextObjectWithMeta;
    const center = getTextCenter(text);
    text.autoLayoutId = uid("layout");
    text.autoLayoutOrder = 1;
    text.autoLayoutDirection = "horizontal";
    text.autoLayoutGap = 16;
    text.autoLayoutAlign = "center";
    text.autoLayoutAnchorX = center.x;
    text.autoLayoutAnchorY = center.y;
    applyAutoLayouts(canvas);
    syncSelection();
    refreshTextObjects();
  }

  function pinAutoLayoutAnchorForActiveText() {
    const canvas = fabricCanvasRef.current;
    const active = canvas?.getActiveObject();
    if (!canvas || !active || !isFabricText(active) || !active.autoLayoutId) return;
    updateAutoLayoutAnchorFromBounds(canvas, active.autoLayoutId);
    applyAutoLayouts(canvas);
    syncSelection();
    refreshTextObjects();
  }

  function updateTypeStyle(fieldType: FieldType, patch: Partial<TypeStyleConfig>) {
    setTypeStyles((current) => {
      const next = mergeTypeStyles(current);
      next[fieldType] = { ...next[fieldType], ...patch };
      return next;
    });
  }

  function applySelectedTypeStyle() {
    const canvas = fabricCanvasRef.current;
    const active = canvas?.getActiveObject();
    if (!canvas || !active || !isFabricText(active)) return;
    const style = getTypeStyle(typeStyles, active.fieldType);
    applyTypeStyleToText(active, style);
    refreshCanvasTextLayout(canvas);
    syncSelection();
    refreshTextObjects();
  }


  function createProjectPayload(): ProjectFile {
    return {
      version: 2,
      savedAt: new Date().toISOString(),
      templates: getTemplatesWithCurrentCanvas(),
      fonts,
      rows,
      excelColumns,
      typeStyles
    };
  }

  async function loadProject(project: ProjectFile) {
    setBusyMessage("Opening project...");
    try {
      await registerFontAssets(project.fonts || []);
      loadedTemplateIdRef.current = null;
      setFonts(project.fonts || []);
      setRows(project.rows || []);
      setExcelColumns(project.excelColumns || []);
      setTemplates(project.templates || []);
      setTypeStyles(mergeTypeStyles(project.typeStyles));
      setActiveTemplateId(project.templates?.[0]?.id || null);
      setPreviewItems([]);
      setSelectedText(null);
    } finally {
      setBusyMessage(null);
    }
  }

  function downloadProject() {
    const project = createProjectPayload();
    const blob = new Blob([JSON.stringify(project, null, 2)], { type: "application/json;charset=utf-8" });
    saveAs(blob, "invitation-project.json");
  }

  function saveProjectToBrowser() {
    const project = createProjectPayload();
    localStorage.setItem(STORAGE_KEY, JSON.stringify(project));
  }

  async function loadProjectFromBrowser() {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    await loadProject(JSON.parse(raw) as ProjectFile);
  }

  async function handleOpenProject(file: File | null) {
    if (!file) return;
    const raw = await file.text();
    await loadProject(JSON.parse(raw) as ProjectFile);
  }

  function downloadSampleCsv() {
    const csv = [
      "Name,Alias,Title",
      "Nguyen Van A,Anh A,CEO",
      "Tran Thi B,Chi B,Marketing Manager",
      "Le Minh C,Minh C,Partner"
    ].join("\n");
    saveAs(new Blob([csv], { type: "text/csv;charset=utf-8" }), "sample-guests.csv");
  }

  async function renderTemplateForRow(
    template: TemplateState,
    row: GuestRow,
    rowIndex: number,
    scale = exportScale,
    quality = jpegQuality
  ): Promise<RenderedOutput> {
    await document.fonts.ready;
    const canvas = await loadTemplateIntoStaticCanvas(template, typeStyles);

    canvas.getObjects().forEach((obj) => {
      if (!isFabricText(obj)) return;
      const column = obj.excelColumn;
      if (column) obj.set("text", displayValue(row[column]));
    });

    refreshCanvasTextLayout(canvas);
    canvas.renderAll();
    const dataUrl = canvas.toDataURL({
      format: "jpeg",
      quality,
      multiplier: scale
    });
    canvas.dispose();

    const guestName = pickGuestName(row, rowIndex);
    const fileName = `${safeFileName(template.name.replace(/\.[^.]+$/, ""))}_${safeFileName(guestName)}.jpg`;
    return {
      fileName,
      dataUrl,
      width: template.width,
      height: template.height
    };
  }

  async function renderOutputs(scope: ExportScope, limit?: number) {
    const currentTemplates = getTemplatesWithCurrentCanvas();
    const selectedTemplates = scope === "all" ? currentTemplates : currentTemplates.filter((t) => t.id === activeTemplateId);
    const renderRows = rows.length ? rows : [{} as GuestRow];
    const cappedRows = typeof limit === "number" ? renderRows.slice(0, limit) : renderRows;
    const outputs: RenderedOutput[] = [];

    for (const template of selectedTemplates) {
      for (let rowIndex = 0; rowIndex < cappedRows.length; rowIndex += 1) {
        setBusyMessage(`Rendering ${outputs.length + 1} / ${selectedTemplates.length * cappedRows.length}...`);
        outputs.push(await renderTemplateForRow(template, cappedRows[rowIndex], rowIndex));
      }
    }

    setBusyMessage(null);
    return outputs;
  }

  async function generatePreview() {
    if (!activeTemplateId) return;
    setPreviewItems([]);
    try {
      const outputs = await renderOutputs(previewScope, previewLimit);
      setPreviewItems(
        outputs.map((output, index) => ({
          id: uid("preview"),
          fileName: output.fileName,
          dataUrl: output.dataUrl,
          rowIndex: index,
          templateName: output.fileName.split("_")[0] || "Template"
        }))
      );
    } finally {
      setBusyMessage(null);
    }
  }

  async function exportJpegs(scope: ExportScope) {
    if (!activeTemplateId) return;
    try {
      const outputs = await renderOutputs(scope);
      if (outputs.length === 0) return;
      if (outputs.length === 1) {
        saveAs(dataUrlToBlob(outputs[0].dataUrl), outputs[0].fileName);
        return;
      }

      const zip = new JSZip();
      outputs.forEach((output) => {
        zip.file(output.fileName, dataUrlToBlob(output.dataUrl));
      });
      setBusyMessage("Creating ZIP...");
      const blob = await zip.generateAsync({
        type: "blob",
        compression: "DEFLATE",
        compressionOptions: { level: 6 }
      });
      saveAs(blob, "invitations-jpeg.zip");
    } finally {
      setBusyMessage(null);
    }
  }

  async function exportPdf(scope: ExportScope) {
    if (!activeTemplateId) return;
    try {
      const outputs = await renderOutputs(scope);
      if (outputs.length === 0) return;
      const first = outputs[0];
      const pdf = new jsPDF({
        orientation: first.width >= first.height ? "landscape" : "portrait",
        unit: "px",
        format: [first.width, first.height],
        compress: true
      });

      outputs.forEach((output, index) => {
        if (index > 0) {
          pdf.addPage([output.width, output.height], output.width >= output.height ? "landscape" : "portrait");
        }
        pdf.addImage(output.dataUrl, "JPEG", 0, 0, output.width, output.height, undefined, "FAST");
      });

      pdf.save(scope === "all" ? "invitations-all.pdf" : "invitations-active.pdf");
    } finally {
      setBusyMessage(null);
    }
  }

  const hasTemplates = templates.length > 0;
  const mappedCount = textObjects.filter((item) => item.excelColumn).length;

  return (
    <main className="app-shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">MVP batch invitation editor</p>
          <h1>Invitation Card Editor</h1>
        </div>
        <div className="topbar-actions">
          <button className="ghost" onClick={saveProjectToBrowser} disabled={!hasTemplates}>
            Save browser
          </button>
          <button className="ghost" onClick={loadProjectFromBrowser}>
            Load browser
          </button>
          <button className="primary" onClick={downloadProject} disabled={!hasTemplates}>
            Download project
          </button>
        </div>
      </header>

      <section className="workspace">
        <aside className="panel left-panel">
          <div className="panel-section">
            <h2>1. Assets</h2>
            <label className="upload-card">
              <input type="file" accept="image/*" multiple onChange={(event) => void handleUploadImages(event.target.files)} />
              <span>Upload card images</span>
              <small>PNG, JPG, WEBP. Multiple files become multiple templates.</small>
            </label>
            <label className="upload-card">
              <input type="file" accept=".ttf,.otf,.woff,.woff2,font/*" multiple onChange={(event) => void handleUploadFont(event.target.files)} />
              <span>Upload custom fonts</span>
              <small>Fonts are embedded in project JSON.</small>
            </label>
            <label className="upload-card">
              <input type="file" accept=".xlsx,.csv" onChange={(event) => void handleUploadSpreadsheet(event.target.files?.[0] || null)} />
              <span>Upload Excel / CSV</span>
              <small>First row must contain column names.</small>
            </label>
            <label className="upload-card compact">
              <input type="file" accept="application/json,.json" onChange={(event) => void handleOpenProject(event.target.files?.[0] || null)} />
              <span>Open project JSON</span>
            </label>
            <button className="ghost full" onClick={downloadSampleCsv}>Download sample CSV</button>
          </div>

          <div className="panel-section">
            <h2>2. Templates</h2>
            {templates.length === 0 ? (
              <p className="muted">Upload at least one invitation image to start.</p>
            ) : (
              <div className="template-list">
                {templates.map((template) => (
                  <button
                    key={template.id}
                    className={template.id === activeTemplateId ? "template-item active" : "template-item"}
                    onClick={() => switchTemplate(template.id)}
                  >
                    <img src={template.imageDataUrl} alt={template.name} />
                    <span>{template.name}</span>
                    <small>{template.width} x {template.height}px</small>
                  </button>
                ))}
              </div>
            )}
          </div>

          <div className="panel-section">
            <h2>3. Text fields</h2>
            <button className="primary full" onClick={addTextBox} disabled={!activeTemplate}>
              Add text field
            </button>
            <div className="text-object-list">
              {textObjects.map((item) => (
                <button key={item.id} className="text-object" onClick={() => selectTextObject(item.id)}>
                  <strong>{item.text || "Untitled"}</strong>
                  <span>
                    {FIELD_LABELS[item.fieldType]} {item.excelColumn ? `-> ${item.excelColumn}` : "-> not mapped"}
                    {item.autoLayoutId ? ` | layout: ${item.autoLayoutId}` : ""}
                  </span>
                </button>
              ))}
            </div>
          </div>

          <div className="panel-section stats">
            <h2>Data status</h2>
            <div><strong>{rows.length}</strong><span>guest rows</span></div>
            <div><strong>{excelColumns.length}</strong><span>columns</span></div>
            <div><strong>{mappedCount}</strong><span>mapped fields</span></div>
            {excelColumns.length > 0 && (
              <p className="column-list">Columns: {excelColumns.join(", ")}</p>
            )}
          </div>
        </aside>

        <section className="center-stage">
          <div className="canvas-toolbar">
            <div>
              <strong>{activeTemplate?.name || "No template selected"}</strong>
              {activeTemplate && <span>{activeTemplate.width} x {activeTemplate.height}px</span>}
            </div>
            <div className="toolbar-actions">
              <label>
                Zoom
                <input type="range" min="0.2" max="1.5" step="0.05" value={zoom} onChange={(event) => setZoom(Number(event.target.value))} />
                <span>{Math.round(zoom * 100)}%</span>
              </label>
            </div>
          </div>

          <div className="canvas-wrap">
            {!hasTemplates && (
              <div className="empty-state">
                <h2>Upload a card image to begin</h2>
                <p>Add text fields, map them to Excel columns, preview, then export JPEG/PDF/ZIP.</p>
              </div>
            )}
            <canvas ref={canvasElementRef} />
          </div>

          <div className="preview-panel">
            <div className="preview-header">
              <div>
                <h2>Batch preview</h2>
                <p>Preview renders with actual Excel data and current template styles.</p>
              </div>
              <div className="preview-controls">
                <select value={previewScope} onChange={(event) => setPreviewScope(event.target.value as ExportScope)}>
                  <option value="active">Active template</option>
                  <option value="all">All templates</option>
                </select>
                <label>
                  Limit
                  <input type="number" min="1" max="200" value={previewLimit} onChange={(event) => setPreviewLimit(Number(event.target.value))} />
                </label>
                <button className="primary" onClick={() => void generatePreview()} disabled={!activeTemplateId}>
                  Generate preview
                </button>
              </div>
            </div>
            {previewItems.length === 0 ? (
              <p className="muted">No preview yet.</p>
            ) : (
              <div className="preview-grid">
                {previewItems.map((item) => (
                  <figure key={item.id}>
                    <img src={item.dataUrl} alt={item.fileName} />
                    <figcaption>{item.fileName}</figcaption>
                  </figure>
                ))}
              </div>
            )}
          </div>
        </section>

        <aside className="panel right-panel">
          <div className="panel-section">
            <h2>Field type styles</h2>
            <p className="muted small-text">Set màu, gradient, stroke, shadow theo từng loại field. Mọi text cùng loại sẽ tự đồng bộ style.</p>
            <div className="type-style-stack">
              {(["name", "alias", "title", "custom"] as FieldType[]).map((fieldType) => {
                const config = typeStyles[fieldType];
                return (
                  <div key={fieldType} className="mini-section">
                    <div className="section-title-row">
                      <div>
                        <strong>{FIELD_LABELS[fieldType]}</strong>
                        <span>{fieldType}</span>
                      </div>
                    </div>
                    <div className="form-grid">
                      <label>
                        Fill type
                        <select value={config.fillMode} onChange={(event) => updateTypeStyle(fieldType, { fillMode: event.target.value as FillMode })}>
                          <option value="solid">Solid</option>
                          <option value="gradient">Gradient</option>
                        </select>
                      </label>
                      {config.fillMode === "solid" ? (
                        <label>
                          Color
                          <input type="color" value={config.fill} onChange={(event) => updateTypeStyle(fieldType, { fill: event.target.value })} />
                        </label>
                      ) : (
                        <>
                          <label>
                            Gradient from
                            <input type="color" value={config.gradientFrom} onChange={(event) => updateTypeStyle(fieldType, { gradientFrom: event.target.value })} />
                          </label>
                          <label>
                            Gradient to
                            <input type="color" value={config.gradientTo} onChange={(event) => updateTypeStyle(fieldType, { gradientTo: event.target.value })} />
                          </label>
                          <label>
                            Angle
                            <input type="number" step="5" value={config.gradientAngle} onChange={(event) => updateTypeStyle(fieldType, { gradientAngle: Number(event.target.value) })} />
                          </label>
                        </>
                      )}

                      <label>
                        Stroke on
                        <input type="checkbox" checked={config.strokeEnabled} onChange={(event) => updateTypeStyle(fieldType, { strokeEnabled: event.target.checked })} />
                      </label>
                      <label>
                        Stroke color
                        <input type="color" value={config.stroke} disabled={!config.strokeEnabled} onChange={(event) => updateTypeStyle(fieldType, { stroke: event.target.value })} />
                      </label>
                      <label>
                        Stroke width
                        <input type="number" min="0" step="0.5" value={config.strokeWidth} disabled={!config.strokeEnabled} onChange={(event) => updateTypeStyle(fieldType, { strokeWidth: Number(event.target.value) })} />
                      </label>

                      <label>
                        Shadow on
                        <input type="checkbox" checked={config.shadowEnabled} onChange={(event) => updateTypeStyle(fieldType, { shadowEnabled: event.target.checked })} />
                      </label>
                      <label>
                        Shadow color
                        <input type="color" value={config.shadowColor} disabled={!config.shadowEnabled} onChange={(event) => updateTypeStyle(fieldType, { shadowColor: event.target.value })} />
                      </label>
                      <label>
                        Blur
                        <input type="number" min="0" step="1" value={config.shadowBlur} disabled={!config.shadowEnabled} onChange={(event) => updateTypeStyle(fieldType, { shadowBlur: Number(event.target.value) })} />
                      </label>
                      <label>
                        Offset X
                        <input type="number" step="1" value={config.shadowOffsetX} disabled={!config.shadowEnabled} onChange={(event) => updateTypeStyle(fieldType, { shadowOffsetX: Number(event.target.value) })} />
                      </label>
                      <label>
                        Offset Y
                        <input type="number" step="1" value={config.shadowOffsetY} disabled={!config.shadowEnabled} onChange={(event) => updateTypeStyle(fieldType, { shadowOffsetY: Number(event.target.value) })} />
                      </label>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          <div className="panel-section">
            <h2>Text properties</h2>
            {!selectedText ? (
              <p className="muted">Select a text field on the canvas to edit its style and data mapping.</p>
            ) : (
              <div className="form-grid">
                <datalist id="auto-layout-groups">
                  {autoLayoutGroups.map((group) => (
                    <option key={group} value={group} />
                  ))}
                </datalist>

                <label className="span-2">
                  Text
                  <textarea value={selectedText.text} onChange={(event) => updateActiveText({ text: event.target.value })} />
                </label>

                <label>
                  Field type
                  <select value={selectedText.fieldType} onChange={(event) => updateActiveText({ fieldType: event.target.value as FieldType })}>
                    <option value="name">Ten</option>
                    <option value="alias">Bi danh</option>
                    <option value="title">Chuc vu</option>
                    <option value="custom">Tuy chinh</option>
                  </select>
                </label>

                <label>
                  Excel column
                  <select value={selectedText.excelColumn} onChange={(event) => updateActiveText({ excelColumn: event.target.value })}>
                    <option value="">Not mapped</option>
                    {excelColumns.map((column) => (
                      <option key={column} value={column}>{column}</option>
                    ))}
                  </select>
                </label>

                <div className="span-2 button-row">
                  <button type="button" className="ghost" onClick={applySelectedTypeStyle}>Apply style of current field type</button>
                </div>

                <label>
                  Font
                  <select value={selectedText.fontFamily} onChange={(event) => updateActiveText({ fontFamily: event.target.value })}>
                    <option value="Arial">Arial</option>
                    <option value="Georgia">Georgia</option>
                    <option value="Times New Roman">Times New Roman</option>
                    <option value="Verdana">Verdana</option>
                    {fonts.map((font) => (
                      <option key={font.id} value={font.family}>{font.fileName}</option>
                    ))}
                  </select>
                </label>

                <label>
                  Size
                  <input type="number" min="4" value={selectedText.fontSize} onChange={(event) => updateActiveText({ fontSize: Number(event.target.value) })} />
                </label>

                <label>
                  Box width
                  <input type="number" min="20" value={selectedText.width} onChange={(event) => updateActiveText({ width: Number(event.target.value) })} />
                </label>

                <label>
                  Opacity
                  <input type="number" min="0" max="1" step="0.05" value={selectedText.opacity} onChange={(event) => updateActiveText({ opacity: Number(event.target.value) })} />
                </label>

                <label>
                  Fill type
                  <select value={selectedText.fillMode} onChange={(event) => updateActiveText({ fillMode: event.target.value as FillMode })}>
                    <option value="solid">Solid color</option>
                    <option value="gradient">Gradient</option>
                  </select>
                </label>

                {selectedText.fillMode === "solid" ? (
                  <label>
                    Color
                    <input type="color" value={selectedText.fill} onChange={(event) => updateActiveText({ fill: event.target.value })} />
                  </label>
                ) : (
                  <>
                    <label>
                      Gradient from
                      <input type="color" value={selectedText.gradientFrom} onChange={(event) => updateActiveText({ gradientFrom: event.target.value })} />
                    </label>
                    <label>
                      Gradient to
                      <input type="color" value={selectedText.gradientTo} onChange={(event) => updateActiveText({ gradientTo: event.target.value })} />
                    </label>
                    <label>
                      Gradient angle
                      <input type="number" step="5" value={selectedText.gradientAngle} onChange={(event) => updateActiveText({ gradientAngle: Number(event.target.value) })} />
                    </label>
                  </>
                )}

                <div className="mini-section span-2">
                  <div className="section-title-row">
                    <div>
                      <strong>Stroke</strong>
                      <span>Outline for selected text</span>
                    </div>
                  </div>
                  <div className="subgrid">
                    <label>
                      Enable stroke
                      <input type="checkbox" checked={selectedText.strokeEnabled} onChange={(event) => updateActiveText({ strokeEnabled: event.target.checked })} />
                    </label>
                    <label>
                      Stroke color
                      <input type="color" value={selectedText.stroke} disabled={!selectedText.strokeEnabled} onChange={(event) => updateActiveText({ stroke: event.target.value })} />
                    </label>
                    <label>
                      Stroke width
                      <input type="number" min="0" step="0.5" value={selectedText.strokeWidth} disabled={!selectedText.strokeEnabled} onChange={(event) => updateActiveText({ strokeWidth: Number(event.target.value) })} />
                    </label>
                  </div>
                </div>

                <div className="mini-section span-2">
                  <div className="section-title-row">
                    <div>
                      <strong>Shadow</strong>
                      <span>Drop shadow for selected text</span>
                    </div>
                  </div>
                  <div className="subgrid">
                    <label>
                      Enable shadow
                      <input type="checkbox" checked={selectedText.shadowEnabled} onChange={(event) => updateActiveText({ shadowEnabled: event.target.checked })} />
                    </label>
                    <label>
                      Shadow color
                      <input type="color" value={selectedText.shadowColor} disabled={!selectedText.shadowEnabled} onChange={(event) => updateActiveText({ shadowColor: event.target.value })} />
                    </label>
                    <label>
                      Blur
                      <input type="number" min="0" step="1" value={selectedText.shadowBlur} disabled={!selectedText.shadowEnabled} onChange={(event) => updateActiveText({ shadowBlur: Number(event.target.value) })} />
                    </label>
                    <label>
                      Offset X
                      <input type="number" step="1" value={selectedText.shadowOffsetX} disabled={!selectedText.shadowEnabled} onChange={(event) => updateActiveText({ shadowOffsetX: Number(event.target.value) })} />
                    </label>
                    <label>
                      Offset Y
                      <input type="number" step="1" value={selectedText.shadowOffsetY} disabled={!selectedText.shadowEnabled} onChange={(event) => updateActiveText({ shadowOffsetY: Number(event.target.value) })} />
                    </label>
                  </div>
                </div>

                <label>
                  X
                  <input type="number" value={selectedText.left} onChange={(event) => updateActiveText({ left: Number(event.target.value) })} />
                </label>

                <label>
                  Y
                  <input type="number" value={selectedText.top} onChange={(event) => updateActiveText({ top: Number(event.target.value) })} />
                </label>

                <label>
                  Angle
                  <input type="number" value={selectedText.angle} onChange={(event) => updateActiveText({ angle: Number(event.target.value) })} />
                </label>

                <label>
                  Auto spacing
                  <select value={selectedText.autoSpacingMode} onChange={(event) => updateActiveText({ autoSpacingMode: event.target.value as AutoSpacingMode })}>
                    <option value="manual">Manual letter spacing</option>
                    <option value="fitWidth">Fit text to box width</option>
                  </select>
                </label>

                <label>
                  Letter space
                  <input
                    type="number"
                    value={selectedText.charSpacing}
                    disabled={selectedText.autoSpacingMode === "fitWidth"}
                    onChange={(event) => updateActiveText({ charSpacing: Number(event.target.value) })}
                  />
                </label>

                {selectedText.autoSpacingMode === "fitWidth" && (
                  <>
                    <label>
                      Min letter space
                      <input type="number" value={selectedText.minCharSpacing} onChange={(event) => updateActiveText({ minCharSpacing: Number(event.target.value) })} />
                    </label>
                    <label>
                      Max letter space
                      <input type="number" value={selectedText.maxCharSpacing} onChange={(event) => updateActiveText({ maxCharSpacing: Number(event.target.value) })} />
                    </label>
                  </>
                )}

                <label>
                  Line height
                  <input type="number" min="0.5" step="0.05" value={selectedText.lineHeight} onChange={(event) => updateActiveText({ lineHeight: Number(event.target.value) })} />
                </label>

                <div className="align-buttons span-2">
                  <button type="button" className={selectedText.textAlign === "left" ? "active" : ""} onClick={() => updateActiveText({ textAlign: "left" })}>Left</button>
                  <button type="button" className={selectedText.textAlign === "center" ? "active" : ""} onClick={() => updateActiveText({ textAlign: "center" })}>Center</button>
                  <button type="button" className={selectedText.textAlign === "right" ? "active" : ""} onClick={() => updateActiveText({ textAlign: "right" })}>Right</button>
                </div>

                <div className="mini-section span-2">
                  <div className="section-title-row">
                    <div>
                      <strong>Auto layout</strong>
                      <span>Use same group ID for fields like “Ông” + name, then set order and gap.</span>
                    </div>
                    <button type="button" className="ghost small-button" onClick={createAutoLayoutGroupForActiveText}>New group</button>
                  </div>

                  <label className="span-2">
                    Group ID
                    <input
                      list="auto-layout-groups"
                      value={selectedText.autoLayoutId}
                      placeholder="Example: honorific-name"
                      onChange={(event) => updateActiveText({ autoLayoutId: event.target.value })}
                    />
                  </label>

                  <div className="subgrid">
                    <label>
                      Order
                      <input type="number" step="1" value={selectedText.autoLayoutOrder} onChange={(event) => updateActiveText({ autoLayoutOrder: Number(event.target.value) })} />
                    </label>
                    <label>
                      Direction
                      <select value={selectedText.autoLayoutDirection} onChange={(event) => updateActiveText({ autoLayoutDirection: event.target.value as AutoLayoutDirection })}>
                        <option value="horizontal">Horizontal</option>
                        <option value="vertical">Vertical</option>
                      </select>
                    </label>
                    <label>
                      Gap
                      <input type="number" value={selectedText.autoLayoutGap} onChange={(event) => updateActiveText({ autoLayoutGap: Number(event.target.value) })} />
                    </label>
                    <label>
                      Group align
                      <select value={selectedText.autoLayoutAlign} onChange={(event) => updateActiveText({ autoLayoutAlign: event.target.value as AutoLayoutAlign })}>
                        <option value="left">Left</option>
                        <option value="center">Center</option>
                        <option value="right">Right</option>
                      </select>
                    </label>
                    <label>
                      Anchor X
                      <input type="number" value={selectedText.autoLayoutAnchorX} onChange={(event) => updateActiveText({ autoLayoutAnchorX: Number(event.target.value) })} />
                    </label>
                    <label>
                      Anchor Y
                      <input type="number" value={selectedText.autoLayoutAnchorY} onChange={(event) => updateActiveText({ autoLayoutAnchorY: Number(event.target.value) })} />
                    </label>
                  </div>

                  <div className="button-row">
                    <button type="button" className="ghost" onClick={pinAutoLayoutAnchorForActiveText} disabled={!selectedText.autoLayoutId}>Pin current position</button>
                    <button type="button" className="ghost" onClick={() => updateActiveText({ autoLayoutId: "" })} disabled={!selectedText.autoLayoutId}>Remove layout</button>
                  </div>
                </div>

                <div className="span-2 two-buttons">
                  <button type="button" className="ghost" onClick={duplicateActiveText}>Duplicate</button>
                  <button type="button" className="danger" onClick={deleteActiveText}>Delete</button>
                </div>
              </div>
            )}
          </div>

          <div className="panel-section">
            <h2>Export</h2>
            <div className="form-grid">
              <label>
                JPEG quality
                <input type="number" min="0.1" max="1" step="0.01" value={jpegQuality} onChange={(event) => setJpegQuality(Number(event.target.value))} />
              </label>
              <label>
                Export scale
                <input type="number" min="1" max="4" step="0.5" value={exportScale} onChange={(event) => setExportScale(Number(event.target.value))} />
              </label>
            </div>
            <p className="muted small-text">Scale 2 or 3 creates high resolution JPEGs and sharper PDFs, but uses more memory.</p>
            <div className="export-buttons">
              <button className="primary" onClick={() => void exportJpegs("active")} disabled={!activeTemplateId}>JPEG active</button>
              <button className="primary" onClick={() => void exportPdf("active")} disabled={!activeTemplateId}>PDF active</button>
              <button className="ghost" onClick={() => void exportJpegs("all")} disabled={!activeTemplateId}>JPEG all / ZIP</button>
              <button className="ghost" onClick={() => void exportPdf("all")} disabled={!activeTemplateId}>PDF all</button>
            </div>
          </div>

          <div className="panel-section help-box">
            <h2>Workflow</h2>
            <ol>
              <li>Upload image templates.</li>
              <li>Upload fonts if needed.</li>
              <li>Add text fields and position them on the card.</li>
              <li>Upload Excel/CSV and map each text field to a column.</li>
              <li>Preview, then export JPEG, PDF, or ZIP.</li>
              <li>Download project JSON to continue later.</li>
            </ol>
          </div>
        </aside>
      </section>

      {busyMessage && <div className="busy-toast">{busyMessage}</div>}
    </main>
  );
}
