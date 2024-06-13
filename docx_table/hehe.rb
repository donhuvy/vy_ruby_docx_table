require 'docx'

doc = Docx::Document.open('foo.docx')

doc.paragraphs.each do |p|
    puts p
end

puts "---------- END :) -------"

doc.bookmarks.each_pair do |bookmark_name, bookmark_object|
    puts bookmark_name
end

# Create a Docx::Document object for our existing docx file
doc = Docx::Document.open('van_table.docx')

first_table = doc.tables[0]
puts first_table.row_count
puts first_table.column_count
puts first_table.rows[0].cells[0].text
puts first_table.columns[0].cells[0].text

# Iterate through tables
doc.tables.each do |table|
    table.rows.each do |row| # Row-based iteration
        row.cells.each do |cell|
            puts cell.text
        end
    end

    table.columns.each do |column| # Column-based iteration
        column.cells.each do |cell|
            puts cell.text
        end
    end
end