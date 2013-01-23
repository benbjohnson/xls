class Xls
  class Selection
    ############################################################################
    #
    # Error Classes
    #
    ############################################################################

    class SelectionFormatError < StandardError; end


    ############################################################################
    #
    # Static Methods
    #
    ############################################################################

    # Parses an Excel style [COLUMN][ROW]:[COLUMN][ROW] format into a
    # selection.
    #
    # @param [String]  str The Excel-style selection.
    #
    # @return [Selection]  A selection object.
    def self.parse(str)
      m, tl_col, tl_row, br_col, br_row = *str.to_s.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/)
      raise SelectionFormatError.new("Invalid selection: #{str}") if m.nil?
      
      # Default bottom-right for single cell selection.
      br_col = tl_col if br_col.nil?
      br_row = tl_row if br_row.nil?
      
      # Convert column letters to numbers.
      tl_col = col_to_index(tl_col)
      br_col = col_to_index(br_col)
      
      # Convert rows to zero-based indices.
      tl_row = tl_row.to_i - 1
      br_row = br_row.to_i - 1
      
      # Invert range if necessary.
      tl_col, br_col = [tl_col, br_col].min, [tl_col, br_col].max
      tl_row, br_row = [tl_row, br_row].min, [tl_row, br_row].max
      
      # Return a selection object.
      return Xls::Selection.new((tl_col..br_col), (tl_row..br_row))
    end

    # Converts column letters to integer indices.
    def self.col_to_index(letters)
      value = 0
      letters.upcase.split('').each_with_index do |letter, index|
        value = value + ((letter.ord - "A".ord) * (26 ** index))
      end
      return value
    end
    

    ############################################################################
    #
    # Constructor
    #
    ############################################################################

    def initialize(columns, rows)
      self.columns = columns
      self.rows = rows
    end


    ############################################################################
    #
    # Attributes
    #
    ############################################################################

    # A range of row indices that the selection covers.
    attr_accessor :rows

    # A range of column indices that the selection covers.
    attr_accessor :columns
    
    # An array of indicies (top-level column, top-left row, bottom-right
    # column, bottom-right row).
    def indices
      return [columns.begin, rows.begin, columns.end, rows.end]
    end
  end
end