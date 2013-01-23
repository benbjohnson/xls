class Xls
  class Enumerator
    ############################################################################
    #
    # Constructor
    #
    ############################################################################

    def initialize(options={})
      self.selection = options[:selection]
      self.procs = options[:procs] || []
    end


    ############################################################################
    #
    # Attributes
    #
    ############################################################################

    # The selection to enumerate over.
    attr_accessor :selection
    
    # A list of procs to run on each cell.
    attr_accessor :procs
    

    ############################################################################
    #
    # Methods
    #
    ############################################################################
    
    # Executes a set of procs on each cell in a selection.
    #
    # @param [Workbook]  input The input workbook.
    def process(input)
      # Loop over each worksheet.
      (0...input.sheet_count).each do |sheet_index|
        sheet = input.worksheet(sheet_index)
        
        # Loop over each row.
        sheet.each do |row|
          next unless selection.nil? || selection.rows.nil? || selection.rows.cover?(row.idx)
          
          # Loop over each cell.
          row.each_with_index do |cell, col_index|
            next unless selection.nil? || selection.columns.nil? || selection.columns.cover?(col_index)
            
            # Run each proc.
            procs.each do |proc|
              proc.call(cell.to_s, col_index, row.idx)
            end
          end
        end
      end
      
      return nil
    end
  end
end