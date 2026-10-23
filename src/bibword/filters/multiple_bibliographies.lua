-- BibWord scoped bibliography renderer.
--
-- Each bibword-bibliography Div is populated from citations in either the
-- whole document or its containing Heading 1, 2, or 3 section. Citeproc is
-- called on isolated scope documents to generate the reference lists, then
-- once on the complete document to render the in-text citations without an
-- additional global bibliography.

local function fail(message)
  error('BibWord multiple bibliographies: ' .. message)
end

local function has_class(element, class_name)
  for _, value in ipairs(element.classes) do
    if value == class_name then
      return true
    end
  end
  return false
end

local function marker_scope(block)
  if block.t ~= 'Div' or not has_class(block, 'bibword-bibliography') then
    return nil
  end
  return block.attributes['data-bibword-scope'] or 'document'
end

local function scope_level(scope)
  if scope == 'document' then
    return nil
  end
  local level = tonumber(scope:match('^heading([123])$'))
  if not level then
    fail('unsupported scope ' .. tostring(scope))
  end
  return level
end

local function scope_bounds(blocks, marker_index, scope)
  local level = scope_level(scope)
  if not level then
    return 1, #blocks
  end
  local first = nil
  for index = marker_index - 1, 1, -1 do
    local block = blocks[index]
    if block.t == 'Header' and block.level <= level then
      if block.level ~= level then
        fail('marker ' .. blocks[marker_index].identifier
          .. ' is not inside a Heading ' .. level .. ' section')
      end
      first = index
      break
    end
  end
  if not first then
    fail('marker ' .. blocks[marker_index].identifier
      .. ' has no preceding Heading ' .. level)
  end
  local last = #blocks
  for index = first + 1, #blocks do
    local block = blocks[index]
    if block.t == 'Header' and block.level <= level then
      last = index - 1
      break
    end
  end
  if marker_index < first or marker_index > last then
    fail('marker ' .. blocks[marker_index].identifier .. ' is outside its heading scope')
  end
  return first, last
end

local function citations_in_scope(blocks, first, last)
  local citations = pandoc.Blocks({})
  for index = first, last do
    if not marker_scope(blocks[index]) then
      blocks[index]:walk({
        Cite = function(cite)
          citations:insert(pandoc.Plain({ pandoc.Cite(cite.content, cite.citations) }))
        end,
      })
    end
  end
  return citations
end

local function generated_references(marker, citation_blocks, meta)
  if #citation_blocks == 0 then
    fail('marker ' .. marker.identifier .. ' has no managed citations in its scope')
  end
  local local_meta = pandoc.Meta(meta)
  local_meta['suppress-bibliography'] = pandoc.MetaBool(false)
  local_meta.citeproc = pandoc.MetaBool(true)
  local refs = pandoc.Div(marker.content, pandoc.Attr('refs'))
  citation_blocks:insert(refs)
  local rendered = pandoc.utils.citeproc(pandoc.Pandoc(citation_blocks, local_meta))
  for _, block in ipairs(rendered.blocks) do
    if block.t == 'Div' and block.identifier == 'refs' then
      block.identifier = marker.identifier
      block.classes:insert('bibword-bibliography')
      block.attributes['data-bibword-id'] = marker.attributes['data-bibword-id']
      block.attributes['data-bibword-scope'] = marker.attributes['data-bibword-scope']
      block = block:walk({
        Div = function(entry)
          if entry.identifier:match('^ref%-') then
            entry.identifier = marker.identifier .. '-' .. entry.identifier
          end
          return entry
        end,
      })
      return block
    end
  end
  fail('citeproc did not generate references for marker ' .. marker.identifier)
end

function Pandoc(doc)
  local replacements = {}
  local marker_count = 0
  for index, block in ipairs(doc.blocks) do
    local scope = marker_scope(block)
    if scope then
      marker_count = marker_count + 1
      local first, last = scope_bounds(doc.blocks, index, scope)
      local citations = citations_in_scope(doc.blocks, first, last)
      replacements[index] = generated_references(block, citations, doc.meta)
    end
  end
  if marker_count == 0 then
    fail('no scoped bibliography markers were found')
  end
  for index, replacement in pairs(replacements) do
    doc.blocks[index] = replacement
  end
  doc.meta['suppress-bibliography'] = pandoc.MetaBool(true)
  -- A citation may belong to several lists, so one HTML anchor cannot identify
  -- its reference unambiguously. Keep the per-list entry IDs unique and disable
  -- citation links until a scope-aware link strategy is introduced.
  doc.meta['link-citations'] = pandoc.MetaBool(false)
  doc.meta.citeproc = pandoc.MetaBool(false)
  return pandoc.utils.citeproc(doc)
end
