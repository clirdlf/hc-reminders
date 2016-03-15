require 'roo'
require 'sequel'
require 'chronic'
require 'mail'

workbook = Roo::Spreadsheet.open('HiddenCollectionsFM.xlsx')
workbook.default_sheet = workbook.sheets[0]

Mail.defaults do
  delivery_method :smtp, address: "localhost", port: 1025
end

headers = Hash.new
workbook.row(1).each_with_index {|header, i| headers[header] = i }

((workbook.first_row + 1)..workbook.last_row).each do |row|
  id           = workbook.row(row)[headers['Grant Contract No.']]
  last_name    = workbook.row(row)[headers['Grants_Contracts LOG::Contact Name Last']]
  first_name   = workbook.row(row)[headers['Grants_Contracts LOG::Contact Name First']]
  email        = workbook.row(row)[headers['Grants_Contracts LOG::Contact Email']]
  organization = workbook.row(row)[headers['Grants_Contracts LOG::Contact Organization']]
  date_end     = workbook.row(row)[headers['Grants_Contracts LOG::Date End']]
  active       = workbook.row(row)[headers['Grants_Contracts LOG::Disposition']]
  final_due    = workbook.row(row)[headers['Grants_Contracts LOG::Final Narr Due']]

  if(active == 'Active')

    unless final_due.nil?
      today = Date.today
      today.to_datetime
      final_due.to_datetime

      difference =  (final_due - today).to_i

      puts "#{final_due} | #{difference}"

      if( difference <= 0)

      end

      email_body = <<-EMAIL
      To: #{email}
      From: hiddencollections@clir.org
      Subject: Hidden Collections Reminder

      Yo, #{first_name}, you need to send in your report that is due on #{final_due}.

      - CLIR
      EMAIL

      #puts email_body


    end

  end



end


DB = Sequel.sqlite

#DB.create_table :grants do

#end
#
#
#


# id => Grant Contract No. (Y)
# Grants_Contracts LOG::Contact Name First (AS)
# Grants_Contracts LOG::Contact Name Last (AT)
# Grants_Contracts LOG::Contact Organization (AU)
# Grants_Contracts LOG::Contact Email (AQ)
# Grants_Contracts LOG::Date End (BC)
# Grants_Contracts LOG::Disposition (BE) looking for active
# Grants_Contracts LOG::Final Narr Due (BG)
#
