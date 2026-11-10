import * as exporter from './markdown-export.mjs';
import apa from '../src/bibdocqmd/styles/apa.csl';
import harvard1 from '../src/bibdocqmd/styles/harvard1.csl';
import filter from '../src/bibdocqmd/filters/multiple_bibliographies.lua';

const assets = { styles: { apa, harvard1 }, filter };
export const prepareMarkdownExport = (bytes, options) => exporter.prepareMarkdownExport(bytes, { ...options, assets });
export const finishMarkdownExport = (preparation, markdown, options) =>
  exporter.finishMarkdownExport(preparation, markdown, { ...options, assets });
