# frozen_string_literal: true

require "test_helper"

class TestOxmlMaker < Minitest::Test
  def test_that_it_has_a_version_number
    refute_nil ::OxmlMaker::VERSION
  end

  def test_module_exists
    assert_kind_of Module, OxmlMaker
  end

  def test_error_class_exists
    assert_kind_of Class, OxmlMaker::Error
    assert OxmlMaker::Error < StandardError
  end
end
