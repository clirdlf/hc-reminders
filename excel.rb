#! /usr/bin/env ruby
# frozen_string_literal: true

require 'erb'
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

##
# Extract data from the Excel cells

def extract_fields(i)
  {
    id:            @workbook.row(i)[@headers['Grant Contract No.']],
    last_name:     @workbook.row(i)[@headers['Grants_Contracts LOG::Contact Name Last']],
    first_name:    @workbook.row(i)[@headers['Grants_Contracts LOG::Contact Name First']],
    email:         @workbook.row(i)[@headers['Grants_Contracts LOG::Contact Email']],
    organization:  @workbook.row(i)[@headers['Grants_Contracts LOG::Contact Organization']],
    project:       @workbook.row(i)[@headers['Grants_Contracts LOG::Description']],
    date_end:      @workbook.row(i)[@headers['Grants_Contracts LOG::Date End']],
    narrative_end: @workbook.row(i)[@headers['Grants_Contracts LOG::Narr 1 Due']],
    active:        @workbook.row(i)[@headers['Grants_Contracts LOG::Disposition']],
    final_due:     @workbook.row(i)[@headers['Grants_Contracts LOG::Final Narr Due']]
  }
end

##
# Calculate the difference between the beginning of the current month and the
# due date

def check_date(date)
  # calculate the first day of the month
  # uses ActiveSupport DateAndTime#Calculations
  # http://apidock.com/rails/DateAndTime/Calculations
  start_of_month = Date.today.beginning_of_month
  (date - start_of_month).to_i
end

##
# Send an email

def narrative_email(row)

  html_template = ERB.new(File.read("templates/30_day_email.html.erb"))
  text_template = ERB.new(File.read("templates/30_day_email.txt.erb"))

   Mail.deliver do
    to (row[:email]).to_s
    from 'Hidden Collections Grant <hiddencollections@clir.org>'
    subject 'Hidden Collections Report Reminder'

    text_part do
      body text_template.result(binding)
    end

    html_part do
      content_type 'text/html; charset=UTF-8'
      body html_template.result(binding)
    end
  end
end

##
# Parse the Excel export

def parse_file
  ((@workbook.first_row + 1)..@workbook.last_row).each do |row|
    row_data = extract_fields(row)

    next unless row_data[:active] == 'Active'
    diff = check_date(row_data[:date_end])

    # 30 - check_date(row_data[:date_end]) # TODO

    if diff >= 25 && diff <= 35
      narrative_email(row_data)
      puts "Sending an email to #{row_data[:first_name]} #{row_data[:last_name]} at #{row_data[:organization]}; their report is due #{row_data[:date_end]}."
    end
  end
end

read_spreadsheet
set_headers
parse_file
