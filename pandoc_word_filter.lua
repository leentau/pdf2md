-- Merge adjacent Word Source Code paragraphs into one fenced code block.
-- Assigning a class makes Pandoc's GFM writer choose a fence instead of four
-- leading spaces, which keeps indentation stable inside lists and callouts.
function Blocks(blocks)
  local result = pandoc.List()
  local pending = nil

  local function flush()
    if pending ~= nil then
      if #pending.classes == 0 then
        pending.classes:insert("text")
      end
      result:insert(pending)
      pending = nil
    end
  end

  for _, block in ipairs(blocks) do
    if block.t == "CodeBlock" then
      if pending == nil then
        pending = block
      else
        pending.text = pending.text .. "\n" .. block.text
      end
    else
      flush()
      result:insert(block)
    end
  end
  flush()
  return result
end
