require 'rexml/document'

class Office365Data
	attr_accessor :batches
	attr_accessor :major_version
	attr_accessor :minor_version
	attr_accessor :description
	attr_accessor :data_interchange_type
	attr_accessor :case_id

	def initialize
		@batches = []
	end

	def total_documents
		result = 0
		@batches.each{|b|result += b.documents.size}
		return result
	end

	def parse_xml_file(xml_file_path)
		xml_file = File.new(xml_file_path)
		xml_doc = REXML::Document.new(xml_file)

		# Collect data from root node
		xml_doc.elements.each("Root") do |root_node|
			@major_version = root_node.attributes["MajorVersion"]
			@minor_version = root_node.attributes["MinorVersion"]
			@description = root_node.attributes["Description"]
			@data_interchange_type = root_node.attributes["DataInterchangeType"]
			@case_id = root_node.attributes["CaseId"]
		end

		# Collect batch data
		xml_doc.elements.each("Root/Batch") do |batch_node|
			@batches << Office365Batch.new(batch_node)
		end
	end

	def each_document(&block)
		@batches.each do |current_batch|
			current_batch.documents.each do |current_document|
				yield(current_document)
			end
		end
	end
end

class Office365Batch
	attr_accessor :documents

	def initialize(batch_node)
		@documents = []
		batch_node.elements.each("Documents/Document") do |document_node|
			@documents << Office365Document.new(document_node)
		end
	end
end

class Office365Document
	attr_accessor :tags
	attr_accessor :files
	attr_accessor :doc_id
	attr_accessor :doc_type
	attr_accessor :mime_type
	attr_accessor :custodian
	attr_accessor :location_uri
	attr_accessor :description

	def initialize(document_node)
		@tags = []
		@files = []
		@doc_id = document_node.attributes["DocID"]
		@doc_type = document_node.attributes["DocType"]
		@mime_type = document_node.attributes["MimeType"]
		# Collect tag nodes
		document_node.elements.each("Tags/Tag") do |tag_node|
			@tags << Office365Tag.new(tag_node)
		end
		# Collect file nodes
		document_node.elements.each("Files/File") do |file_node|
			@files << Office365File.new(file_node)
		end
		# Collect other
		document_node.elements.each("Locations/Location/Custodian") do |custodian_node|
			@custodian = custodian_node.text
		end
		document_node.elements.each("Locations/Location/LocationURI") do |location_uri_node|
			@location_uri = location_uri_node.text
		end
		document_node.elements.each("Locations/Location/Description") do |description_node|
			@description = description_node.text
		end
	end
end

class Office365Tag
	attr_accessor :tag_name
	attr_accessor :tag_data_type
	attr_accessor :tag_value

	def initialize(tag_node)
		@tag_name = tag_node.attributes["TagName"]
		@tag_data_type = tag_node.attributes["TagDataType"]
		@tag_value = tag_node.attributes["TagValue"]
	end
end

class Office365File
	attr_accessor :file_type
	attr_accessor :file_path
	attr_accessor :file_name
	attr_accessor :file_hash
	attr_accessor :file_hash_type

	def initialize(file_node)
		@file_type = file_node.attributes["FileType"]
		file_node.elements.each("ExternalFile") do |external_file_node|
			@file_path = external_file_node.attributes["FilePath"]
			@file_name = external_file_node.attributes["FileName"]
			hash_attribute = external_file_node.attributes["Hash"]
			hash_parts = hash_attribute.split(/:/)
			@file_hash_type = hash_parts[0]

			# Conver hash types to Nuix equivalent field
			@file_hash_type = "sha-256" if @file_hash_type == "SHA256"
			@file_hash_type = "sha-1" if @file_hash_type == "SHA1"

			@file_hash = hash_parts[1]
		end
	end
end