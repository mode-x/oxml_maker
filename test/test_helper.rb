# frozen_string_literal: true

$LOAD_PATH.unshift File.expand_path("../lib", __dir__)
require "oxml_maker"

require "minitest/autorun"
require "ostruct"
require "fileutils"
require "tmpdir"

# Helper method to create test objects
def create_test_object(attributes)
  OpenStruct.new(attributes)
end
