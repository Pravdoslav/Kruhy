function Div(block)
  if FORMAT:match("typst") and block.classes:includes("word-note") then
    local content = { pandoc.RawBlock("typst", "#pad(left: 1.25cm)[") }
    for _, child in ipairs(block.content) do
      table.insert(content, child)
    end
    table.insert(content, pandoc.RawBlock("typst", "]"))
    return content
  end
end
