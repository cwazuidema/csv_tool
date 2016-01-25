require "spreadsheet"
require "HighLine"
require "colorize"

@cli = HighLine.new

def find_all_xls_files
# find all xls files in the current folder		
	puts "Yoyo, I've found these files bro:".blue
	books = Dir['*.xls']
	books.each do |book|
		puts book.red
	end
	open_xls
end

def open_xls
	book = Spreadsheet.open @cli.ask "Which book do you want to open? (use full filename inc .xls)".blue
	sheet1 = book.worksheet 0
	options_for_book(book, sheet1)
end

def options_for_book(book, sheet1)
	answer = @cli.ask "Type 'headers' to view the headers or 'client' to enter a new order".blue
  see_headers(sheet1) if answer == 'headers'
  add_client(book, sheet1) if answer == 'client'
end

def see_headers(sheet1)
	header = sheet1.row(0)
  header.each do |header|
 	  puts header.red
  end
end

def add_client(book, sheet1)
	header = sheet1.row(0)
	header.each_with_index do |header,index|
		sheet1.rows[1][index] = @cli.ask "What data do we need for #{ header.red }"
  end
  book.write @cli.ask "What is the filename of this file (add .xls at the end)"
end



find_all_xls_files