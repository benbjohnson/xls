require 'test_helper'

class TestEnumerator < MiniTest::Unit::TestCase
  def setup
    @enumerator = Xls::Enumerator.new
  end
  
  ######################################
  # Execute
  ######################################
  
  def test_map_enumerate
    workbook = Spreadsheet.open('fixtures/enumerator/basic.xls')
    @enumerator.action = :map
    @enumerator.proc = "text.to_s + '!'"
    @enumerator.process(workbook)
    assert_worksheet [
      ['a!', 'b!', 'c!'],
      ['d!', 'e!', 'f!'],
      ['g!', 'h!', 'i!'],
      ],
      workbook.worksheet(0)
  end

  def test_select_enumerate
    workbook = Spreadsheet.open('fixtures/enumerator/basic.xls')
    @enumerator.action = :select
    @enumerator.proc = "text.to_s.ord % 2 == 0"
    @enumerator.process(workbook)
    assert_worksheet [
      ['', 'b', ''],
      ['d', '', 'f'],
      ['', 'h', ''],
      ],
      workbook.worksheet(0)
  end

  def test_reject_enumerate
    workbook = Spreadsheet.open('fixtures/enumerator/basic.xls')
    @enumerator.action = :reject
    @enumerator.proc = "text.to_s.ord % 2 == 0"
    @enumerator.process(workbook)
    assert_worksheet [
      ['a', '', 'c'],
      ['', 'e', ''],
      ['g', '', 'i'],
      ],
      workbook.worksheet(0)
  end
end
