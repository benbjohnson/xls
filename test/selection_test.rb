require 'test_helper'

class TestSelection < MiniTest::Unit::TestCase
  ######################################
  # Parsing
  ######################################
  
  def test_parse_top_left_cell
    assert_equal [0, 0, 0, 0], Xls::Selection.parse("A1").indices
    assert_equal [0, 0, 0, 0], Xls::Selection.parse("A1:A1").indices
  end

  def test_parse_single_cell
    assert_equal [2, 7, 2, 7], Xls::Selection.parse("C8").indices
  end

  def test_parse_range
    assert_equal [2, 7, 4, 9], Xls::Selection.parse("C8:E10").indices
  end

  def test_parse_inverse_range
    assert_equal [2, 7, 4, 9], Xls::Selection.parse("E10:C8").indices
    assert_equal [2, 7, 4, 9], Xls::Selection.parse("C10:E8").indices
  end
end
