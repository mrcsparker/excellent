'use strict';

type XmlJsonPrimitive = string | null;
type XmlJsonObject = {
  [key: string]: XmlJsonValue | undefined;
};
type XmlJsonValue = XmlJsonPrimitive | XmlJsonObject | XmlJsonValue[];

function xmlToJson(xml: Node): XmlJsonValue {
  let object: XmlJsonObject | XmlJsonPrimitive = {};

  if (xml.nodeType === 1) {
    const element = xml as Element;

    if (element.attributes.length > 0) {
      const attributesNode: XmlJsonObject = {};

      object['@'] = attributesNode;

      for (let index = 0; index < element.attributes.length; index += 1) {
        const attribute = element.attributes.item(index);

        if (attribute === null) {
          continue;
        }

        attributesNode[attribute.nodeName] = attribute.nodeValue;
      }
    }
  } else if (xml.nodeType === 3) {
    object = xml.nodeValue;
  }

  if (!xml.hasChildNodes()) {
    return object;
  }

  if (typeof object !== 'object' || object === null) {
    return object;
  }

  for (let index = 0; index < xml.childNodes.length; index += 1) {
    const item = xml.childNodes.item(index);
    const existingNode = object[item.nodeName];

    if (existingNode === undefined) {
      object[item.nodeName] = xmlToJson(item);
      continue;
    }

    const normalizedNode = Array.isArray(existingNode) ? existingNode : [existingNode];

    normalizedNode.push(xmlToJson(item));
    object[item.nodeName] = normalizedNode;
  }

  return object;
}

function fromBase26(value: string): number {
  const normalizedValue = value.toUpperCase();
  let decimalValue = 0;

  if (normalizedValue === null || normalizedValue === undefined || normalizedValue.length === 0) {
    return -1;
  }

  for (let index = 0; index < normalizedValue.length; index += 1) {
    const characterIndex = normalizedValue.charCodeAt(normalizedValue.length - index - 1) - 'A'.charCodeAt(0);
    decimalValue += (26 ** index) * (characterIndex + 1);
  }

  return decimalValue - 1;
}

function toBase26(value: number): string {
  let remainingValue = Math.abs(value);
  let converted = '';
  let hasIterated = false;

  do {
    let remainder = remainingValue % 26;

    if (hasIterated && remainingValue < 25) {
      remainder -= 1;
    }

    converted = String.fromCharCode(remainder + 'A'.charCodeAt(0)) + converted;
    remainingValue = Math.floor((remainingValue - remainder) / 26);
    hasIterated = true;
  } while (remainingValue > 0);

  return converted;
}

function isNumber(value: unknown): boolean {
  return !Number.isNaN(parseFloat(String(value))) && Number.isFinite(Number(value));
}

function getRowFromCell(value: string): number {
  const rowMatch = value.match(/[0-9]+/gi);

  if (rowMatch === null || rowMatch[0] === undefined) {
    throw new Error('Invalid cell reference: ' + value);
  }

  return parseInt(rowMatch[0], 10) - 1;
}

function getColFromCell(value: string): number {
  const columnMatch = value.match(/[A-Z]+/gi);

  if (columnMatch === null || columnMatch[0] === undefined) {
    throw new Error('Invalid cell reference: ' + value);
  }

  const normalizedValue = columnMatch[0];
  return fromBase26(normalizedValue);
}

function each<T>(
  collection: T[] | Record<string, T> | null | undefined,
  iterator: (value: T, keyOrIndex: string | number, source: T[] | Record<string, T>) => void,
  context: unknown = undefined
): void {
  if (collection === null || collection === undefined) {
    return;
  }

  if (Array.isArray(collection)) {
    collection.forEach(function eachArrayItem(value, index) {
      iterator.call(context, value, index, collection);
    });
    return;
  }

  for (const [key, currentValue] of Object.entries(collection)) {
    iterator.call(context, currentValue, key, collection);
  }
}

const Util = {
  each,
  fromBase26,
  getColFromCell,
  getRowFromCell,
  isNumber,
  toBase26,
  xmlToJson
};

export {
  Util,
  each,
  fromBase26,
  getColFromCell,
  getRowFromCell,
  isNumber,
  toBase26,
  type XmlJsonObject,
  type XmlJsonPrimitive,
  type XmlJsonValue,
  xmlToJson
};
