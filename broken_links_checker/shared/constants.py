import os

FILE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
print(FILE_DIR)
HOME_URL = 'https://www.mindef.gov.sg/web/wcm/connect/mindef/mindef-content/home'
CATEGORIES = 'categories'
PARAM_CATEGORY = 'param_category'
PARAM_PAGE = 'param_page'
SITE_AREA_NAME = {
    'news-and-events': {
        'latest-releases': {
            CATEGORIES: ['news-releases', 'speeches', 'parliamentary-statements', 'forum-letter-replies', 'replies-to-media-queries', 'clarification-of-issues', 'others'],
            PARAM_CATEGORY: 'selectedCategories',
            PARAM_PAGE: 'wcm_page.MENU-latest-releases',
        },
        'events-and-advisories': {
            CATEGORIES: [''],
            PARAM_CATEGORY: '',
            PARAM_PAGE: 'wcm_page.MENU-events-and-advisories',
        },
    },
}
URL_PARAM_SITEAREANAME = 'siteAreaName'
URL_PARAMS = {
    URL_PARAM_SITEAREANAME: 'mindef-content/home/{l1}/{l2}',
    'srv': 'cmpnt',
    'cmpntid': 'dcb39e68-0637-4383-b587-29be9bb9bea5',
    'source': 'library',
    'cache': 'none',
    'contentcache': 'none',
    'connectorcache': 'none',
}