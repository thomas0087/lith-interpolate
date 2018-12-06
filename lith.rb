require 'rubygems'
require 'rubyXL'

# Helper methods
def parse_row(row, offset)
  {
    row_number: row.index_in_collection,
    data_set:   row[6].value,
    hole_id:    row[offset].value,
    depth_from: row[offset + 1].value,
    depth_to:   row[offset + 2].value,
    lith_plot:  row[offset + 3]&.value,
    lith1_code: row[offset + 4].value
  }
end

def add_row(sheet, row_number, measurement)
  sheet.add_cell(row_number, 0, measurement[:data_set])
  sheet.add_cell(row_number, 1, measurement[:hole_id])
  sheet.add_cell(row_number, 2, measurement[:depth_from])
  sheet.add_cell(row_number, 3, measurement[:depth_to])
  sheet.add_cell(row_number, 4, measurement[:lith_plot])
  sheet.add_cell(row_number, 5, measurement[:lith1_code])
end

def split_row(source, options = {})
  options = options.merge(row_number: nil)
  source.dup.merge(options)
end
# End helper methods

# Load the spreadsheet
excel_file = 'LithPlotConversionOD.xlsx'
workbook = RubyXL::Parser.parse(excel_file)
raw_lith_sheet = workbook['To_Convert_to_Lith_Plot']

# Read in existing (origin values)
holes = []
measurements = []
row_offset = 7

raw_lith_sheet.each do |row|
  next if row[row_offset].nil? || row[row_offset].value == 'Hole_ID'
  holes << row[row_offset].value
  measurements << parse_row(row, row_offset)
end

holes = holes.uniq

puts "* Imported #{holes.count} holes"
puts "* Imported #{measurements.count} measurements"

# Read in updated values
updated_measurements = []
row_offset = 12
raw_lith_sheet.each do |row|
  next if row[row_offset].nil?
  updated_measurements << parse_row(row, row_offset)
end

puts "* Imported #{updated_measurements.count} updated measurements"

# Insert the updated values in the correct location
updated_measurements.each do |update|
  # get all measurements from that hole
  hole_measurements = measurements.select { |m| m[:hole_id] == update[:hole_id] }

  # update matches existing depth
  exact_match = hole_measurements.select { |m| m[:depth_from] == update[:depth_from] && m[:depth_to] == update[:depth_to] }
  if exact_match.any?
    exact_match.first[:lith_plot] = update[:lith1_code]
    next
  end

  # get measurements in the depth range of the update
  range_matches = hole_measurements.select { |m| m[:depth_from] < update[:depth_to] && m[:depth_to] > update[:depth_from] }
  if range_matches.any?
    #puts range_matches.count

    # order matches by depth_from, increasing
    range_matches = range_matches.sort_by { |m| m[:depth_from] }

    # check the leading match
    first_m = range_matches.first
    if first_m[:depth_from] == update[:depth_from]
      # if it has the same starting point we just update the lith code
      first_m[:lith_plot] = update[:lith1_code]
    else
      # if there is a different starting point we need to split the value
      #   create new measurement from m[:depth_from] to update[:depth_from] leaving values as is
      measurements << split_row(first_m, { depth_to: update[:depth_from] })
      #   create new measurement from update[:depth_from] to m[:depth_to] with lith_plot = update[:lith1_code]
      measurements << split_row(first_m, { depth_from: update[:depth_from], lith_plot: update[:lith1_code] })

      # remove first_interval (can match on row_number)
      measurements.delete_if { |m| m[:row_number] == first_m[:row_number] }
    end

    # if there's only a single range match our work here is done
    # i.e. fully inside
    next if range_matches.count == 1

    # check the trailing interval
    last_m = range_matches.last
    if last_m[:depth_to] == update[:depth_to]
      # if it has the same ending point we just update the lith code
      last_m[:lith_plot] = update[:lith1_code]
    else
      # if there is a different ending point we need to split the value
      #   create new measurement from m[:depth_from] to update[:depth_to] with lith_plot = update[:lith1_code]
      measurements << split_row(last_m, { depth_to: update[:depth_to], lith_plot: update[:lith1_code] })
      #   create new measurement from update[:depth_to] to m[:depth_to] leaving values as is
      measurements << split_row(last_m, { depth_from: update[:depth_to] })

      # remove last interval (can match on row_number)
      measurements.delete_if { |m| m[:row_number] == last_m[:row_number] }
    end

    # if there's only two matches our work here is done
    next if range_matches.count == 2

    # for the rest of the intervals lith_plot = update[lith1_code]
    range_matches[1..-2].each do |m|
      m[:lith_plot] = update[:lith1_code]
    end

    next
  end

  # otherwise the new measurement is deeper than any existing one, append it
  # I don't think this is likely but should serve as a catchall
  measurements << update
end

puts "resultant measurements = #{measurements.count}"

# Create the results sheet
processed_lith_sheet = workbook.add_worksheet('Processed Lith')
row_number = 0

# Add header row
processed_lith_sheet.add_cell(row_number, 0, 'DataSet')
processed_lith_sheet.add_cell(row_number, 1, 'Hole_ID')
processed_lith_sheet.add_cell(row_number, 2, 'Depth_From')
processed_lith_sheet.add_cell(row_number, 3, 'Depth_To')
processed_lith_sheet.add_cell(row_number, 4, 'Lith_Plot')
processed_lith_sheet.add_cell(row_number, 5, 'Lith1_Code')

# for each hole get all the measurements and concat them in depth order
holes.each do |hole|
  measurements.select { |m| m[:hole_id] == hole }.sort_by { |m| m[:depth_from] }.each do |measurement|
    row_number += 1
    add_row(processed_lith_sheet, row_number, measurement)
  end
end

workbook.write("LithPlotConversionTD.xlsx")

puts "Inserted #{row_number} new rows"

puts '*** Processing complete ***'
