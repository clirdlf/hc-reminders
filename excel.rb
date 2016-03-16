#! /usr/bin/env ruby
# frozen_string_literal: true

require 'optparse'
require 'pp'

require 'active_support'
require 'active_support/core_ext'
require 'mail'
require 'roo'

##
# Script for generating emails for Hidden Collections Grantees
#
# Author:: Wayne Graham (mailto:wgraham@clir.org)
# Copyright:: Copyright (c) 2016 Council on Library and Information Resources
# License:: MIT

class HiddenCollectionsMailer
  attr_reader :workbook

  def initialize(file)
  end
end

class Parser
  Version = '0.0.1'.freeze

  class ScriptOptions
    attr_accessor :file, :verbose, :dry_run

    def initialize
      self.file = 'HiddenCollectionsFM.xslx'
      self.verbose = false
      self.dry_run = true
    end
  end

  #
  # Return a structure describing the options
  #
  def self.parse(args)
    # The options specified on the command line will be collected in
    # *options*.

    @options = ScriptOptions.new
    option_parser.parse!(args)
    @options
  end

  attr_reader :parser, :options

  def option_parser
    @parser ||= OptionParser.new do |parser|
      parser.banner = 'Usage: excel.rb [options]'
      parser.separator ''
      parser.separator 'Specific options:'

      # add additional options
      set_file_path
      boolean_verbose_option
      set_dry_run

      parser.separator ''
      parser.separator 'Common options:'

      parser.on_tail('-h', '--help', 'Show help message') do
        puts parser
        exit
      end

      parser.on_tail('--version', 'Show Version') do
        puts Version
        exit
      end
    end
  end

  def set_file_path
    parser.on('-f', '--file [FILE PATH]', 'Set the path to the FileMaker Pro Excel export') do |f|
      options.file = f
    end
  end

  def boolean_verbose_option
    parser.on('-v', '--[no]verbose', 'Run verbosely') do |v|
      options.verbose = v
    end
  end

  def set_dry_run
    parser.on('-d', '--[no]dryrun', 'Do a dry run') do |d|
      options.dry_run = d
    end
  end
end

# options = Parser.parse(ARGV)
# pp options

Mail.defaults do
  delivery_method :smtp, address: "localhost", port: 1025
end

##
# Read the export for the Hidden Collections Filemaker Database
#

def read_spreadsheet(file = 'HiddenCollectionsFM.xlsx')
  @workbook = Roo::Spreadsheet.open(file)
  @workbook.default_sheet = @workbook.sheets[0]
end

##
# Read the first row of the spreadsheet and store headers and their index

def set_headers
  @headers = {}
  @workbook.row(1).each_with_index { |header, i| @headers[header] = i }
end

def find_data(i)
  {
    id: @workbook.row(i)[@headers['Grant Contract No.']],
    last_name: @workbook.row(i)[@headers['Grants_Contracts LOG::Contact Name Last']],
    first_name: @workbook.row(i)[@headers['Grants_Contracts LOG::Contact Name First']],
    email: @workbook.row(i)[@headers['Grants_Contracts LOG::Contact Email']],
    organization: @workbook.row(i)[@headers['Grants_Contracts LOG::Contact Organization']],
    date_end: @workbook.row(i)[@headers['Grants_Contracts LOG::Date End']],
    active: @workbook.row(i)[@headers['Grants_Contracts LOG::Disposition']],
    final_due: @workbook.row(i)[@headers['Grants_Contracts LOG::Final Narr Due']]
  }
end

def check_date(date)
  # calculate the first day of the month
  # uses ActiveSupport DateAndTime#Calculations
  # http://apidock.com/rails/DateAndTime/Calculations
  start_of_month = Date.today.beginning_of_month
  difference = (date - start_of_month).to_i
end

def send_email(row)
  mail = Mail.deliver do
    to (row[:email]).to_s
    from 'Hidden Collections Grant <hiddencollections@clir.org>'
    subject 'Hidden Collections Report Reminder'

    text_part do
      body "This is the plain text part\n\nDear #{row[:first_name]},\n\nThis is just a reminder that the report for your Hidden Collection grant is due #{row[:date_end]}"
    end

    html_part do
      content_type 'text/html; charset=UTF-8'
      body "<h1>This is the HTML part</h1><p>Dear #{row[:first_name]},</p><p>This is just a reminder that the report for your Hidden Collections grant is due #{row[:date_end]}</p>"
    end
  end
end

def parse_file
  ((@workbook.first_row + 1)..@workbook.last_row).each do |row|
    row_data = find_data(row)

    next unless row_data[:active] == 'Active'
    diff = check_date(row_data[:date_end])

    if diff >= 85 && diff <= 95
      send_email(row_data)
      puts row_data
    end
    # puts row_data
  end
end

read_spreadsheet
set_headers
parse_file
