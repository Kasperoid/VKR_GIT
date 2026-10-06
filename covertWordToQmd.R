rmarkdown::pandoc_convert(
  input = "ВКР_Устинов.docx",
  to = "markdown",
  output = "ВКР_Устинов.qmd",
  options = c("--extract-media=media", "--wrap=none")
)