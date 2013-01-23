require 'bundler/setup'
require 'minitest/autorun'
require 'mocha'
require 'unindentable'
require 'xls'

class MiniTest::Unit::TestCase
  def assert_worksheet exp, worksheet, msg = nil
    act = []
    worksheet.each do |row|
      act << row.map {|cell| cell.to_s.strip }
    end
    assert_equal(exp, act, msg)
  end
end

