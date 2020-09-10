script_directory = File.dirname(__FILE__)
require File.join(script_directory,"Nx.jar")
java_import "com.nuix.nx.NuixConnection"
java_import "com.nuix.nx.LookAndFeelHelper"
java_import "com.nuix.nx.dialogs.ChoiceDialog"
java_import "com.nuix.nx.dialogs.TabbedCustomDialog"
java_import "com.nuix.nx.dialogs.CommonDialogs"
java_import "com.nuix.nx.dialogs.ProgressDialog"

LookAndFeelHelper.setWindowsIfMetal
NuixConnection.setUtilities($utilities)
NuixConnection.setCurrentNuixVersion(NUIX_VERSION)

load File.join(script_directory,"Office365XmlObjects.rb")

dialog = TabbedCustomDialog.new("Associate Office365 Data")

main_tab = dialog.addTab("main_tab","Main")
main_tab.appendOpenFileChooser("xml_file","XML File","Xtensible Markup Language","xml")
main_tab.appendCheckBox("update_custodian","Update Custodian",true)
main_tab.appendCheckBox("tag_updated_items","Tag Updated Items",true)
main_tab.appendTextField("update_tag_name","Tag Name","Office365Associated")
main_tab.enabledOnlyWhenChecked("update_tag_name","tag_updated_items")

dialog.validateBeforeClosing do |values|
	if !java.io.File.new(values["xml_file"]).exists
		CommonDialogs.showError("Please select a valid XML file.")
		next false
	end

	if CommonDialogs.getConfirmation("The script needs to close all workbench tabs, proceed?") == false
		next false
	end
	next true
end

dialog.display
if dialog.getDialogResult == true
	$window.closeAllTabs
	values = dialog.toMap
	update_custodian = values["update_custodian"]
	tag_updated_items = values["tag_updated_items"]
	update_tag_name = values["update_tag_name"]
	annotater = $utilities.getBulkAnnotater

	ProgressDialog.forBlock do |pd|
		pd.setMainStatusAndLogIt("Parsing XML file...")
		xml_data = Office365Data.new
		xml_data.parse_xml_file(values["xml_file"])
		pd.setMainStatusAndLogIt("Applying data to case...")
		pd.setMainProgress(0,xml_data.total_documents)
		document_index = 0
		all_matched_items = []
		xml_data.each_document do |document|
			pd.setMainProgress(document_index+1)

			# Possible for document node to have no file sub-nodes which means
			# we won't have a hash to find matching Nuix items so we need to
			# check for this and handle appropriately
			if document.files.nil? || document.files.size < 1
				pd.logMessage("No File sub-nodes for Document with ID #{document.doc_id}")
			end

			file_hash = document.files.first.file_hash
			file_hash_type = document.files.first.file_hash_type.downcase
			items = $current_case.search("#{file_hash_type}:\"#{file_hash}\"")
			if items.size < 1
				pd.logMessage("No items with #{file_hash_type}: #{file_hash}")
			else
				all_matched_items += items
				# Carry "tag" data over as custom metadata
				document.tags.each do |tag|
					case tag.tag_data_type
					when "Boolean"
						annotater.putCustomMetadata(tag.tag_name,tag.tag_value == "True",items,nil)
					else
						annotater.putCustomMetadata(tag.tag_name,tag.tag_value || "",items,nil)
					end
				end

				# Carry other data over as custom metadata
				annotater.putCustomMetadata("Custodian",document.custodian || "",items,nil)
				annotater.putCustomMetadata("LocationURI",document.location_uri || "",items,nil)
				annotater.putCustomMetadata("Description",document.description || "",items,nil)
				annotater.putCustomMetadata("DocID",document.doc_id || "",items,nil)
				annotater.putCustomMetadata("DocType",document.doc_type || "",items,nil)
				annotater.putCustomMetadata("MimeType",document.mime_type || "",items,nil)

				if update_custodian
					if !document.custodian.nil?
						annotater.assignCustodian(document.custodian,items)
					else
						pd.logMessage("No custodian value for Document with ID #{document.doc_id}")
					end
				end
			end
			document_index += 1
		end

		if tag_updated_items
			annotater.addTag(update_tag_name,all_matched_items)
		end

		annotater.putCustomMetadata("CaseId",xml_data.case_id || "",all_matched_items,nil)

		pd.setMainStatusAndLogIt("Completed")
	end
end

