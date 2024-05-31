class Locators():
    # search_button_xpath = "//button[contains(@data-element, 'search-button')]"
    # searchbar_xpath = "//input[contains(@name, 'q')]"
    # dropdown_xpath = "//select[contains(@name, 's')]"
    # next_results_xpath = '//div[@class="search-results-module-next-page"]'
    # articles_xpath = '//ul[@class="search-results-module-results-menu"]/li'
    # article_title_xpath = './/div/h3[@class="promo-title"]'
    # article_description_xpath = './/p[@class="promo-description"]'
    # article_date_xpath = './/p[@class="promo-timestamp"]'
    # article_image_xpath = './/img[@class="image"]'
    search_button_xpath = '//div[@class="search-button"]'
    # searchbar_xpath = '//input[@class="search-page-input"]'
    # searchbar_xpath = "//input[contains(@name, 'q')]"
    searchbar_xpath = '//*[@id="search"]/input'

    dropdown_xpath = "//select[contains(@name, 's')]"
    category_menu_xpath = '//div[@class="SearchFilter-heading"]'

    # spotlight_xpath = '/html/body/div[4]/bsp-search-results-module/form/div[2]/div/bsp-search-filters/div/main/div[2]'
    individual_article_xpath = './/div[@class="PagePromo-content"]'
    articles_xpath = '//div[@class="SearchResultsModule-results"]//div[@class="PageList-items-item"]'

    article_title_xpath = './/div[@class="PagePromo-title"]/a/span[@class="PagePromoContentIcons-text"]'
    article_description_xpath = './/div[@class="PagePromo-description"]/a/span[@class="PagePromoContentIcons-text"]'
    article_image_xpath = './/div[@class="PagePromo-media"]/a/picture/img[@class="Image"]/@src'
    article_date_xpath = './/span[contains(@class, "Timestamp-template")]'

    # article_title_xpath = '//div[@class="PageList-items"]//bsp-custom-headline//a/span[@class="PagePromoContentIcons-text"]'
    # article_title_xpath = './/span[@class="PagePromoContentIcons-text"]'
    # article_title_xpath = '//div[@class="PagePromo-title"]/a/span[@class="PagePromoContentIcons-text"]'

    # article_description_xpath = '//div[contains(@class, "PagePromo-description")]/a/span[contains(@class, "PagePromoContentIcons-text")]'
    article_description_xpath = './/span[contains(@class, "PagePromoContentIcons-text")]'
    article_date_xpath = '//span[contains(@class, "Timestamp-template")]'
    article_image_xpath = './/img[@class="Image"]'
    next_results_xpath = '//div[@class="Pagination-nextPage"]'
