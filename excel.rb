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
def find_data(i)
  {
    id:           @workbook.row(i)[@headers['Grant Contract No.']],
    last_name:    @workbook.row(i)[@headers['Grants_Contracts LOG::Contact Name Last']],
    first_name:   @workbook.row(i)[@headers['Grants_Contracts LOG::Contact Name First']],
    email:        @workbook.row(i)[@headers['Grants_Contracts LOG::Contact Email']],
    organization: @workbook.row(i)[@headers['Grants_Contracts LOG::Contact Organization']],
    project:      @workbook.row(i)[@headers['Grants_Contracts LOG::Description']],
    date_end:     @workbook.row(i)[@headers['Grants_Contracts LOG::Date End']],
    active:       @workbook.row(i)[@headers['Grants_Contracts LOG::Disposition']],
    final_due:    @workbook.row(i)[@headers['Grants_Contracts LOG::Final Narr Due']]
  }
end

#
# Calculate the difference between the beginning of the current month and the
# due date

def check_date(date)
  # calculate the first day of the month
  # uses ActiveSupport DateAndTime#Calculations
  # http://apidock.com/rails/DateAndTime/Calculations
  start_of_month = Date.today.beginning_of_month
  difference = (date - start_of_month).to_i
end

#
# Send an email
#

def send_email(row)
  mail = Mail.deliver do
    to (row[:email]).to_s
    from 'Hidden Collections Grant <hiddencollections@clir.org>'
    subject 'Hidden Collections Report Reminder'

    text_part do
      body "Dear #{row[:first_name]},

Report Due: #{row[:final_due]}
Grant Ref. No.: #{row[:id]}
Project: #{row[:project]}

Please accept this message as a friendly reminder that a report is due shortly for the Hidden Collections grant referenced above.

In narrative and financial reports, CLIR expects grantees to describe and evaluate the activities achieved during the grant period and to explain how the grant funds were used for those activities. Reports must be submitted via the program's online report form in accordance with our guidelines, which can be found on our Web site:

http://www.clir.org/hiddencollections/recipients/recipients.html#report

To assist you with your financial report, you may find a mock-up of an exemplary financial report through this link (PDF):

http://www.clir.org/hiddencollections/recipients/hidden-collections-financial-report-mockup

If you have any questions about any aspect of the grant, please contact Program Officer Nicole Ferraiolo at nferraiolo@clir.org. We look forward to reading your report.

Best,
Adam

--
Adam Leader-Smith
Program Associate
Council on Library and Information Resources
1707 L St. NW, Suite 650
Washington, D.C., 20036
(202) 939-4759
http://www.clir.org"
    end

    html_part do
      content_type 'text/html; charset=UTF-8'
      body "<p>Dear #{row[:first_name]},</p>

      <p>
        Report Due: <strong>#{row[:final_due]}</strong>
      <br>
        Grant Ref. No.: #{row[:id]}
      <br>
        Project: #{row[:project]}
      </p>

      <p>
        Please accept this message as a friendly reminder that a report is due shortly for the Hidden Collections grant referenced above.
      </p>
      <p>
        In narrative and financial reports, CLIR expects grantees to describe and evaluate the activities achieved during the grant period and to explain how the grant funds were used for those activities. Reports must be submitted via the program's online report
        form in accordance with our guidelines, which can be found on our Web site:
      </p>
      <p>
        <a href='http://www.clir.org/hiddencollections/recipients/recipients.html#report'>http://www.clir.org/hiddencollections/recipients/recipients.html#report</a>
      </p>
      <p>
        To assist you with your financial report, you may find a mock-up of an exemplary financial report through this link (PDF):
      </p>
      <p>
        <a href='http://www.clir.org/hiddencollections/recipients/hidden-collections-financial-report-mockup'>http://www.clir.org/hiddencollections/recipients/hidden-collections-financial-report-mockup</a>
      </p>
      <p>
        If you have any questions about any aspect of the grant, please contact Program Officer Nicole Ferraiolo at <a href='mailto:nferraiolo@clir.org'>nferraiolo@clir.org</a>. We look forward to reading your report.
      </p>
      <p>
        Best,
        <br /> Adam
      </p>
      <p>
        --
        <br> Adam Leader-Smith
        <br> Program Associate
        <br> Council on Library and Information Resources
        <br> 1707 L St. NW, Suite 650
        <br> Washington, D.C., 20036
        <br> (202) 939-4759
        <br> http://www.clir.org
      </p>
"
    end
  end
end

#
# Parse the Excel export
#

def parse_file
  ((@workbook.first_row + 1)..@workbook.last_row).each do |row|
    row_data = find_data(row)

    next unless row_data[:active] == 'Active'
    diff = check_date(row_data[:date_end])

    if diff >= 85 && diff <= 95
      send_email(row_data)
      puts "Sending an email to #{row_data[:first_name]} #{row_data[:last_name]} at #{row_data[:organization]}; their report is due #{row_data[:date_end]}."
    end
  end
end

read_spreadsheet
set_headers
parse_file
