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

  def test_parse_columns_only
    assert_equal [2, nil, 2, nil], Xls::Selection.parse("C").indices
    assert_equal [2, nil, 4, nil], Xls::Selection.parse("C:E").indices
  end

  def test_parse_rows_only
    assert_equal [nil, 2, nil, 2], Xls::Selection.parse("3").indices
    assert_equal [nil, 4, nil, 7], Xls::Selection.parse("5:8").indices
  end
end
