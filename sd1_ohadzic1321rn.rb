# sd1_ohadzic1321rn
require 'google_drive'

class Tabela
include Enumerable
  def initialize(row, col, endRow, endCol)
    pomoc1
    pomoc2
    @row, @col, @endRow, @endCol = row, col, endRow, endCol
    ignore_total_subtotal_rows
  end
  def pomoc1
    @session = GoogleDrive::Session.from_config('config.json')
  end
  def pomoc2
    @ws = @session.spreadsheet_by_key('1Oa1kbt7xiYzQgJrPwh90vMjktAShm7cAZYVTUX1pih4').worksheets[0]
  end
  attr_accessor :row, :col, :endRow, :endCol, :ws, :session
  def ws(value)
    @ws=value
  end

  def ignore_total_subtotal_rows
    @ws.rows.each_with_index do |row, index|
      row_values = row.map(&:downcase)
      if row_values.include?('total') || row_values.include?('subtotal')
        puts "Ignoring row #{index + 1}: #{row_values.join(', ')}"
        next
      end
      test(@row, @col)
    end
  end

  def neki_each
    p "ovo je foreach za tabelu"
    (@row..@endRow).each do |row|
      (@col..@endCol).each do |col|
        p @ws[row, col]
      end
    end
  end
  def [](value)
    trazena_kolona = 0
    (@col..@endCol).each do |col|
      helper = @ws[@row, col]
      if helper==value
        trazena_kolona=col
        break
      end
    end
    if trazena_kolona.nil?
      p "nema kolone"
    else
      (@row..@endRow).each do |row|
        p @ws[row, trazena_kolona]
      end
    end
  end

  def test(startRow, startCol)
    header = @ws.rows[startRow-1][startCol-1..-1]
    header.each_with_index do |header, index|
      define_singleton_method("#{header.downcase.gsub(' ', '_')}") do
        ColumnAccessor.new(@ws, startRow+1, startCol+index)
      end
    end
  end
  class ColumnAccessor
    def initialize(worksheet, start_row, column_index)
      @worksheet = worksheet
      @start_row = start_row
      @column_index = column_index
    end
  
    def sum
      @worksheet.rows[@start_row..-1].transpose[@column_index-1].inject(0) { |sum, value| sum + value.to_f }
    end
  
    def avg
      values = @worksheet.rows[@start_row..-1].transpose[@column_index-1]
      sum = values.inject(0) { |sum, value| sum + value.to_f }
      sum / values.length
    end

    def map(&block)
      @worksheet.rows[@start_row..-1].transpose[@column_index-1].map{ |value| block.call(value.to_i) }
    end
    def select(&block)
      @worksheet.rows[@start_row..-1].transpose[@column_index-1].select{ |value| block.call(value.to_i) }
    end
    def reduce(&block)
      @worksheet.rows[@start_row..-1].transpose[@column_index-1].reduce{ |value, value2| block.call(value.to_i, value2.to_i) }
    end
  end
  def get_table_data(start_row, start_column, end_row, end_column)
    data = []
    (start_row..end_row).each do |row|
      row_data = []
      (start_column..end_column).each do |column|
        if @ws[row, column].to_f != nil
          row_data << @ws[row, column].to_f
        end
      end
      data << row_data
    end

    data
  end
  def set_table_data(data, start_row, start_column)
    data.each_with_index do |row_data, i|
      row_data.each_with_index do |value, j|
        @ws[start_row + i, start_column + j] = value
      end
    end
  end


  def row(value)
    p "ovo je #{value} red tabele:"
    (@col..@endCol).each do |col|
      p @ws[value+7, col]
    end
  end
end




t = Tabela.new(6,4,9,6)
t2 = Tabela.new(6,4,9,6)
t2.ws = t2.instance_variable_get(:@session).spreadsheet_by_key('1Oa1kbt7xiYzQgJrPwh90vMjktAShm7cAZYVTUX1pih4').worksheets[1]

rezt = t.get_table_data(t.instance_variable_get(:@row),t.instance_variable_get(:@col),t.instance_variable_get(:@endRow),
t.instance_variable_get(:@endCol)).zip(t2.get_table_data(t2.instance_variable_get(:@row),t2.instance_variable_get(:@col),
t2.instance_variable_get(:@endRow),t2.instance_variable_get(:@endCol))).map { 
  |row1, row2| row1.zip(row2).map { |cell1, cell2| cell1 + cell2 } }
rezt2 = t.get_table_data(t.instance_variable_get(:@row),t.instance_variable_get(:@col),t.instance_variable_get(:@endRow),
t.instance_variable_get(:@endCol)).zip(t2.get_table_data(t2.instance_variable_get(:@row),t2.instance_variable_get(:@col),
t2.instance_variable_get(:@endRow),t2.instance_variable_get(:@endCol))).map { 
  |row1, row2| row1.zip(row2).map { |cell1, cell2| cell1 - cell2 } }
t3 = Tabela.new(6,4,9,6)
t3.ws = t3.instance_variable_get(:@session).spreadsheet_by_key('1Oa1kbt7xiYzQgJrPwh90vMjktAShm7cAZYVTUX1pih4').worksheets[2]
t3.set_table_data(rezt, t3.instance_variable_get(:@row), t3.instance_variable_get(:@col))
t4 = Tabela.new(6,4,9,6)
t4.ws = t4.instance_variable_get(:@session).spreadsheet_by_key('1Oa1kbt7xiYzQgJrPwh90vMjktAShm7cAZYVTUX1pih4').worksheets[3]
t4.set_table_data(rezt2, t4.instance_variable_get(:@row), t3.instance_variable_get(:@col))
p "sabiranje 2 tabele (obe tabele su iste kao primer iz fajla Skript Domaći)" 
p t3.instance_variable_get(:@ws).rows
p "oduzimanje 2 tabele" 
p t4.instance_variable_get(:@ws).rows
t.row(1)
p t.neki_each
p "ovo je treca kolona:"
p t["Treca kolona"]
#ddx = t["Prva kolona"][6] => ne radi
p "suma prve kolone"
p t.prva_kolona.sum
p "prosek prve kolone"
p t.prva_kolona.avg
p "suma druge kolone"
p t.druga_kolona.sum
p "prosek druge kolone"
p t.druga_kolona.avg
p "suma trece kolone"
p t.treca_kolona.sum
p "prosek trece kolone"
p t.treca_kolona.avg
p "map funkcija"
p t.prva_kolona.map {|cell| cell+=1}
p "select funkcija"
p t.prva_kolona.select {|cell| cell.even?}
p "reduce funkcija"
p t.prva_kolona.reduce {|sum, cell| sum + cell}



t.instance_variable_get(:@ws).save
t.instance_variable_get(:@ws).reload