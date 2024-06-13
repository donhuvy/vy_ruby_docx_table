# require 'docx'
#
# # Retrieve and display paragraphs as html
# doc = Docx::Document.open('minhthu.docx')
# doc.paragraphs.each do |p|
#     puts p.to_html
# end

require 'docx'

# Open the Word document
doc = Docx::Document.open('minhthu.docx')

# Initialize an empty string to hold the HTML content
html_content = "<html>\n<head>\n<title>Document</title>\n</head>\n<body>\n"

# Append each paragraph as HTML
doc.paragraphs.each do |p|
    html_content += "<p>#{p.to_html}</p>\n"
end

# Close the HTML tags
html_content += "</body>\n</html>"

# Write the HTML content to a new file
File.open('minhthu.html', 'w') do |file|
    file.write(html_content)
end

puts "HTML file has been saved as 'minhthu.html'."

