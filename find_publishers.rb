#!/usr/bin/env ruby

# a script to parse an xlsx file and find publisher based on the ISBN

# input: xlsx spreadsheet with the list of ISBNs in column C
# output: an updated spreadsheet with a new column that correspond to publishers

# usage: ruby find_publishers.rb path_to_file

# references:
# http://webservices.amazon.com/scratchpad/index.html

require 'time'
require 'uri'
require 'openssl'
require 'base64'
require 'nokogiri'
require 'open-uri'
require 'rubyXL'


# Your AWS Access Key ID, as taken from the AWS Your Account page
AWS_ACCESS_KEY_ID = ENV['AWS_ACCESS_KEY_ID']
# Your AWS Secret Key corresponding to the above ID, as taken from the AWS Your Account page
AWS_SECRET_KEY = ENV['AWS_SECRET_KEY']
# YourAWS associate tag, taken from the Associates Profile
AssociateTag = ENV['AssociateTag']
# The region you are interested in
ENDPOINT_US = "webservices.amazon.com"
ENDPOINT_CAN = "webservices.amazon.ca"
CALLS_PER_SECOND = 1																		
REQUEST_URI = "/onca/xml"

OUTPUT_FILE_NAME = "/spreadsheet_final_test.xlsx"


def parameters_lookup(isbn)																																																																																										
	params = {
	  "Service" => "AWSECommerceService",
	  "Operation" => "ItemLookup",
	  "AWSAccessKeyId" => AWS_ACCESS_KEY_ID,
	  "AssociateTag" => AssociateTag,
	  "ItemId" => isbn,
	  "IdType" => "ISBN",
	  "ResponseGroup" => "ItemAttributes",
	  "SearchIndex" => "Books"
	}
	return params
end


def create_url(params, endpoint)

	# Set current timestamp if not set
	params["Timestamp"] = Time.now.gmtime.iso8601 if !params.key?("Timestamp")

	# Generate the canonical query
	canonical_query_string = params.sort.collect do |key, value|
	  [URI.escape(key.to_s, Regexp.new("[^#{URI::PATTERN::UNRESERVED}]")), URI.escape(value.to_s, Regexp.new("[^#{URI::PATTERN::UNRESERVED}]"))].join('=')
	end.join('&')

	# Generate the string to be signed
	string_to_sign = "GET\n#{endpoint}\n#{REQUEST_URI}\n#{canonical_query_string}"

	# Generate the signature required by the Product Advertising API
	signature = Base64.encode64(OpenSSL::HMAC.digest(OpenSSL::Digest.new('sha256'), AWS_SECRET_KEY, string_to_sign)).strip()

	# Generate the signed URL
	request_url = "http://#{endpoint}#{REQUEST_URI}?#{canonical_query_string}&Signature=#{URI.escape(signature, Regexp.new("[^#{URI::PATTERN::UNRESERVED}]"))}"

	return request_url
end

def get_response(request_url)
	tries ||= 3	
	request_time = Time.now
	response = open(request_url)	
	rescue OpenURI::HTTPError => error
		while Time.now - request_time < 1 
			# wait until one second passes
		end
		retry unless (tries -= 1).zero?	

		response = error.io	
		puts response.status
		puts response.string
		puts request_url
		return response	
	else
		return response
end		

def prepare_XML(response)
	if response.status[0] != '200'
		puts response.status
		return false, response.status[1]
	end

	doc = Nokogiri::XML(response)
	doc.remove_namespaces!  # TODO improve
	number_results = doc.xpath("//Item").length
	if number_results == 0
		puts "found 0 results"	
		return false, "found 0 results"
	end	
	return true, doc
end


def parse_XML(doc, key)
	list = doc.xpath("//#{key}")
	if list.length == 0
		return 'NOTHING FOUND'
	else
		return list[0].content
	end
end

def update_excel(path_to_file)
	path_to_file = path_to_file
	workbook = RubyXL::Parser.parse path_to_file
	worksheet = workbook[0]
	worksheet = workbook.worksheets[0]
	worksheet.each_with_index do |row, index|

		next if index == 0
		break if index == 4
		isbn = row[2].value.to_s
		title = row[0].value
		author = row[4].value

		puts index

		params = parameters_lookup(isbn)
		request = create_url(params, ENDPOINT_CAN)
		response = get_response(request)
		status, document = prepare_XML(response)

		if not status
			print document
			next
		end

		publisher = parse_XML(document, 'Publisher')
		link = parse_XML(document, 'DetailPageURL')
		title_found = parse_XML(document, 'Title') 

		puts publisher

		worksheet.add_cell(index, 10, publisher)
		worksheet.add_cell(index, 12, title)		
		worksheet.add_cell(index, 14, link)

	end
	dir = File.dirname(path_to_file)
	workbook.write(dir + OUTPUT_FILE_NAME)
end




path_to_file = ARGV[0]
# path_to_file = '/home/tnosov/projects/amazon_books/spreadsheet.xlsx'

update_excel(path_to_file)