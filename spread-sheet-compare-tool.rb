require 'spreadsheet'

workbook = Spreadsheet.open './source/source.xls'
source_sheet_name = 'Sheet1'
salesforce_sheet_name = 'Sheet2'


source_sheet = workbook.worksheet source_sheet_name
salesforce_source_sheet = workbook.worksheet salesforce_sheet_name
source_header_row = source_sheet.row(0)

source_id_col_record_mapping = Hash.new

salesforce_source_id_col_record_mapping = Hash.new
salesforce_header_row = salesforce_source_sheet.row(0)

destination_header_row = Array.new
destination_id_col_record_mapping = Hash.new


##################################### METHODS ############################################

def checkForInteger(val)
  begin
    n = Integer(val)
    return true
  rescue ArgumentError
    return false
  end
end

#################################### METHODS END #########################################
 

######################## Reading Source XLS ########################

source_sheet.each 1 do |row|
  source_col_header_record_mapping = Hash.new
  source_header_row.each_with_index do |col, index|
    if !col.nil?
      source_col_header_record_mapping[col] = row[index];
    end
  end
  source_id_col_record_mapping[row[0]] = source_col_header_record_mapping;
end

###################### Reading Source XLS End #######################


######################## Reading Salesforce XLS ########################

salesforce_source_sheet.each 1 do |row|
  salesforce_source_col_header_record_mapping = Hash.new
  if !row.empty?
    source_header_row.each_with_index do |col, index|
      if !col.nil?
        next if !salesforce_header_row.include? col
        if !row[index].nil?
          salesforce_source_col_header_record_mapping[col] = row[index];
        end
      end
    end
    salesforce_source_id_col_record_mapping[row[0]] = salesforce_source_col_header_record_mapping;
  end
end

###################### Reading Salesforce XLS End #######################


################## Creating Destination XLS Header ######################

destination_header_row.push(source_header_row[0]);

################ Creating Destination XLS Header End #####################



################# Processing Source and Salesforce Data for Comparision ######################


source_id_col_record_mapping.each do |key, value|
  salesforce_row = salesforce_source_id_col_record_mapping[key]
  destination_col_header_record_mapping = Hash.new
  source_header_row.each_with_index  do |col, index|
    next if index == 0
    if !col.nil?
      #salesforce data map
      #salesforce_source_col_header_record_mapping[col] = row[index];
      
      ## don't require to add simple header in destination
      #if !destination_header_row.include? col
      #  destination_header_row.push(col);
      #end
      
      next if !salesforce_row.key?(col)
      
      destination_source_header = "source_#{col}"
      destination_salesforce_header = "salesforce_#{col}"
      
      ## need to add two header one with source prefix
      if !destination_header_row.include? destination_source_header
        destination_header_row.push(destination_source_header);
      end
      
      ## need to add two header one with salesforce prefix
      if !destination_header_row.include? destination_salesforce_header
        destination_header_row.push(destination_salesforce_header);
      end
      destination_col_header_record_mapping[destination_source_header] = value[col];
      destination_col_header_record_mapping[destination_salesforce_header] = salesforce_row[col];
      
      ## need to add difference header
      destination_source_header_differnce = "difference#{col}"
      if !destination_header_row.include? destination_source_header_differnce
        destination_header_row.push(destination_source_header_differnce);
      end
      
      if checkForInteger(value[col])
        destination_col_header_record_mapping[destination_source_header_differnce] = value[col] - salesforce_row[col]
      else
        difference_value = "Matched"
        if !value[col].eql?(salesforce_row[col])
          difference_value = "Value not matched Source Value is #{value[col]} and Salesforce Value is #{salesforce_row[col]}"
        end
        destination_col_header_record_mapping[destination_source_header_differnce] = difference_value
      end
    end
  end
  destination_id_col_record_mapping[key] = destination_col_header_record_mapping
end

############### Processing Source and Salesforce Data for Comparision End ####################



############################# Creating Destination Excel File ###########################

destination_workbook = Spreadsheet::Workbook.new
destination_worksheet = destination_workbook.create_worksheet
destination_worksheet.insert_row 0, destination_header_row

id_field_for_destination = destination_header_row[0];

destination_id_col_record_mapping.each_with_index do |(key, value), index|
  xls_value_to_insert = Array.new
  xls_value_to_insert.push(key)
  destination_header_row.each_with_index  do |record, i|
    next if i == 0
    xls_value_to_insert.push(value[record])
  end
  destination_worksheet.insert_row index+1, xls_value_to_insert
end

destination_workbook.write 'destination_workbook.xls'


########################### Creating Destination Excel File End #########################
