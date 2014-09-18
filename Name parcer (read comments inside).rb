# file parcer by kmar

#this script will parce "month-start.xls" and "month-end.xls" files for names
#it will remove dups in both files first
#it will create 3 pages in a different file, "month-difference.xls"
#page 1 will have new entrys: names that appear in -end, but not in -start
#page 2 will have fired names: thay appear in -start, but not in -end
#page 3 will have updates: names that appear in both places, but have different attributes

require "roo"
require "spreadsheet"

#first we read both files:
print "loading files... "
#startxls = Spreadsheet.open("month-start.xls")
#endxls = Spreadsheet.open("month-end.xls")

startxls = Roo::Excel.new("month-start.xls")
endxls   = Roo::Excel.new("month-end.xls"  )

startxls.default_sheet = startxls.sheets.first
endxls.default_sheet = endxls.sheets.first

puts "DONE"

#sheet1 = startxls.worksheet 0
#sheet2 = endxls.worksheet 0

#now we start comparing:
diffxls = Spreadsheet::Workbook.new
sheet3 = diffxls.create_worksheet :name => "Новые сотрудники" #new
keycol = 1
lastsheetrow = 0

diffrows = []
#new
print "Finding new people... "
for i in (endxls.first_row)..(endxls.last_row) do
  flag = true
  for o in (startxls.first_row)..(startxls.last_row) do
    if endxls.cell(i,keycol) == startxls.cell(o,keycol)
	  flag = false
	end
  end
  
  if flag
	for o in 1..10 do      
	  sheet3[lastsheetrow,o-1] = endxls.cell(i,o)
	end
	lastsheetrow += 1
  end  
end
puts "DONE"

#lastsheetrow += 10
lastsheetrow = 0
#fired
print "Finding fired people... "
sheet4 = diffxls.create_worksheet :name => "Уволенные сотрудники"#fired
#sheet4 = sheet3
for i in (startxls.first_row)..(startxls.last_row) do
  flag = true
  for o in (endxls.first_row)..(endxls.last_row) do    	
    if startxls.cell(i,keycol) == endxls.cell(o,keycol)
	  flag = false	
	end
  end
  
  if flag
	for o in 1..10 do      
	  sheet4[lastsheetrow,o-1] = startxls.cell(i,o)	  
	end
	lastsheetrow += 1
  end  
  
end

puts "DONE"
#differences
sheet5 = diffxls.create_worksheet :name => "Перемещения сотрудников"#some difference
#this is the hard one - we need to find all that don't appear in new or fired (number in front not unique), but do have a difference
print "Finding modifications... "
lastsheetrow = 0
for i in (endxls.first_row)..(endxls.last_row) do
  flag = false
  for o in (startxls.first_row)..(startxls.last_row) do    	
    if endxls.cell(i,keycol) == startxls.cell(o,keycol)
	  # puts "flag = true: #{endxls.cell(i,keycol)} == #{startxls.cell(o,keycol)}"
	  for u in 1..10 do
	    if endxls.cell(i,u) != startxls.cell(o,u)
	      flag = true	
		end
	  end
	end
  end
  
  if flag
	for o in 1..10 do      
	  sheet5[lastsheetrow,o-1] = endxls.cell(i,o)	  
	end
	lastsheetrow += 1
  end  
  
end

puts "DONE"

print "Saving..."
diffxls.write "month-difference.xls"
puts "DONE"










