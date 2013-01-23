require 'test_helper'

class TestColumnizer < MiniTest::Unit::TestCase
  def setup
    @columnizer = Xls::Columnizer.new
  end
  
  ######################################
  # Execute
  ######################################
  
  def test_basic_columnize
    input = Spreadsheet.open('fixtures/columnizer/basic.xls')
    output = @columnizer.process(input)
    assert_worksheet [
      ['KEY', 'VALUE'],
      ['col1', 'a'],
      ['col2', 'b'],
      ['col3', 'c'],
      ['col4', 'd'],
      ['col5', 'e'],
      ['col6', 'f'],
      ['col1', 'g'],
      ['col2', 'h'],
      ['col3', 'i'],
      ['col4', 'j'],
      ['col5', 'k'],
      ['col6', 'l']
      ],
      output.worksheet(0)
  end
end
