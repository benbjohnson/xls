class Xls
  class Enumerator
    ############################################################################
    #
    # Constructor
    #
    ############################################################################

    def initialize(options={})
      self.selection = options[:selection]
      self.proc = options[:proc]
    end


    ############################################################################
    #
    # Attributes
    #
    ############################################################################

    # The action to perform on each cell of data (:none, :map, :select, :reject).
    attr_accessor :action

    # The selection to enumerate over.
    attr_accessor :selection
    
    # The proc to run on each cell.
    attr_accessor :proc

    # A flag stating if headers should be used.
    attr_accessor :use_headers


    ############################################################################
    #
    # Methods
    #
    ############################################################################
    
    # Executes a proc on each cell in a selection.
    #
    # @param [Workbook]  input The input workbook.
    def process(input)
      # Loop over each worksheet.
      (0...input.sheet_count).each do |sheet_index|
        sheet = input.worksheet(sheet_index)
        headers = sheet.row(0)
        
        # Loop over each row.
        sheet.each(use_headers ? 1 : 0) do |row|
          next unless selection.nil? || selection.rows.nil? || selection.rows.cover?(row.idx)
          
          # Loop over each cell.
          row.each_with_index do |cell, col_index|
            next unless selection.nil? || selection.columns.nil? || selection.columns.cover?(col_index)
            
            # Run proc and retrieve the results.
            header = (use_headers ? headers[col_index].to_s : '')
            result = cell_binding(cell.to_s, col_index, row.idx, header).eval(proc)
            
            # Perform an action if one is specified.
            case action
            when :map then row[col_index] = result.to_s
            when :select then row[col_index] = '' unless result
            when :reject then row[col_index] = '' if result
            end
          end
        end
      end
      
      return nil
    end
    
    
    ####################################
    # Bindings
    ####################################

    # Returns the context in which cell procs are run.
    def cell_binding(text, col, row, header)
      return binding
    end
  end
end