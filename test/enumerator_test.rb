require 'test_helper'

class TestEnumerator < MiniTest::Unit::TestCase
  def setup
    @enumerator = Xls::Enumerator.new
  end
  
  ######################################
  # Execute
  ######################################
  
  def test_enumerate
    my_arr = []
    input = Spreadsheet.open('fixtures/enumerator/basic.xls')
    @enumerator.procs = [
      lambda {|cell, col, row| my_arr << cell.to_s},
      lambda {|cell, col, row| my_arr << '0'},
    ]
    @enumerator.process(input)
    assert_equal ["a", "0", "b", "0", "c", "0", "d", "0", "e", "0", "f", "0", "g", "0", "h", "0", "i", "0"], my_arr
  end

  def test_enumerate_with_selection
    my_arr = []
    input = Spreadsheet.open('fixtures/enumerator/basic.xls')
    @enumerator.procs = [lambda {|cell, col, row| my_arr << cell.to_s}]
    @enumerator.selection = Xls::Selection.parse("B2:C3")
    @enumerator.process(input)
    assert_equal ['e', 'f', 'h', 'i'], my_arr
  end
end
