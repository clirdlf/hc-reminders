# -*- coding: UTF-8 -*-
#! /usr/bin/env ruby

# SEE https://github.com/mech/filemaker-ruby#README
# You do have to have the Web Publishing Engine (XML Publishing) enabled

server = Filemaker::Server.new do |config|
  config.host = ENV['FILEMAKER_HOST']
  config.account_name = ENV['FILEMAKER_ACCOUNT_NAME']
  config.password     = ENV['FILEMAKER_PASSWORD']
  config.ssl          = { verify: false }
  config.log          = :curl
end

puts server.databases.all



