class Xls
  class Columnizer
    ############################################################################
    #
    # Constructor
    #
    ############################################################################

    def initialize(options={})
      self.fixed_columns = options[:fixed_columns] || []
    end


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
        output_sheet.row(0).replace(fixed_columns.clone.concat(['KEY', 'VALUE']))

        # Process data.
        column_range = (0...input_headers.length)
        input_sheet.each(1) do |input_row|
          # Determine fixed column values.
          fixed_column_values = fixed_columns.map { '' }
          column_range.each do |index|
            fixed_column_index = fixed_columns.index(input_headers[index])
            fixed_column_values[fixed_column_index] = input_row[index] unless fixed_column_index.nil?
          end
          
          # Write key/value data.
          column_range.each do |index|
            next if fixed_columns.index(input_headers[index])
            output_row = output_sheet.row(output_index)
            output_row.replace(fixed_column_values)

            # Write header and data.
            output_row.push(input_headers[index])
            output_row.push(input_row[index])
    
            output_index = output_index + 1
          end
        end

      end
      
      return output
    end
  end
end