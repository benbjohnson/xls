class Xls
  class Columnizer
    ############################################################################
    #
    # Constructor
    #
    ############################################################################


    ############################################################################
    #
    # Attributes
    #
    ############################################################################

    # A list of names of the columns that should stay fixed.
    attr_accessor :fixed_columns
    

    ############################################################################
    #
    # Methods
    #
    ############################################################################
    
    # Converts the columns on the worksheets of the input workbook from
    # n-column tables to 2-column tables.
    #
    # @param [Workbook]  input The input workbook.
    #
    # @return [Workbook]  The columnized workbook.
    def process(input)
      output = Spreadsheet::Workbook.new()
      
      (0...input.sheet_count).each do |input_sheet_index|
        input_sheet = input.worksheet(input_sheet_index)
        output_sheet = output.create_worksheet()
        output_index = 1

        # Process headers.
        input_headers = input_sheet.row(0).map { |cell| cell.to_s }
        output_sheet.row(0).replace(['KEY', 'VALUE'])

        input_sheet.each(1) do |row|
          (0...input_headers.length).each do |index|
            output_row = output_sheet.row(output_index)

            # Write header and data.
            output_row[0] = input_headers[index]
            output_row[1] = row[index]
    
            output_index = output_index + 1
          end
        end

      end
      
      return output
    end
  end
end