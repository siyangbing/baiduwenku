click_xpath_dict = {
    'load_more_page': ["//body//div[@class='foldpagewg-text']", "//body//div[@class='pagerwg-button']"],
    'end_page': ["//body//div[@class='pagerwg-button' and @style='display: none;']" ],
    'login_cancel_xpath': ["//div[@id='wui-messagebox-cancel-1']"],
}

file_attribute_xpath = {
    "file_name": ["//head/title/text()"],
    "file_type": ["//body//div[contains(@class,'file-type-icon file-icon ')]/@class"],
    'page_data': ["//body//div[@id]/following-sibling::script[1]/text()"]}

content_dict = {
    'doc': ["//body//div[@class='reader-container']//p/text()"],
    'pdf': [],
    'ppt': []}

next_page_cancel_dict = {'doc': ["//div[@id='wui-messagebox-cancel-1']"],
                         'pdf': [],
                         'ppt': []}
