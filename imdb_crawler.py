from typing import List
import requests
from bs4 import BeautifulSoup
import xlwt
import re
import json
import argparse

def get_movie_reviews_cell(url, sheet: xlwt.Worksheet, movie_id: int, cnt:int, is_first=False):

    print('url = ', url)

    res = requests.get(url)
    res.encoding = 'utf-8'
    soup = BeautifulSoup(res.text, 'lxml')

    for review_id, item in enumerate(soup.select(".lister-item-content")):

        title = item.select(".title")[0].text.strip()

        author = item.select(".display-name-link")[0].text

        date = item.select(".review-date")[0].text

        votetext = item.select(".text-muted")[0].text

        upvote = re.findall(r"\d+",votetext)[0]

        totalvote = re.findall(r"\d+", votetext)[1]

        rating = item.select("span.rating-other-user-rating > span")
        if len(rating) == 2:
            rating = rating[0].text
        else:
            rating = ""

        review = item.select(".text")[0].text

        row = [review_id, movie_id, title, author, date, upvote, totalvote, rating, review]
        for i in range(0, len(row)):
            sheet.write(cnt + movie_id * MAX_REVIEW_CNT, i, row[i])

        cnt += 1

        if cnt > MAX_REVIEW_CNT: break

    load_more = soup.find('div', class_='load-more-data')
    url = None
    if is_first:
        if load_more.has_attr('data-key'):
            ajaxurl = load_more['data-ajaxurl']
            base_url = 'https://www.imdb.com/' + ajaxurl + "?ref_=undefined&paginationKey="
            key = load_more['data-key']
            url: str = base_url + key
    else:
        if load_more.has_attr('data-key'):
            key = load_more['data-key']
            url: str = 'https://www.imdb.com/' + key

    return sheet, url, cnt

def get_movie_reviews(url, sheet: xlwt.Worksheet, movie_id: int):

    cnt = 1

    sheet, url, cnt = get_movie_reviews_cell(url, sheet, movie_id, cnt, is_first=True)

    while cnt <= MAX_REVIEW_CNT and url is not None:
        sheet, url, cnt = get_movie_reviews_cell(url, sheet, movie_id, cnt)

    print('Finish movie reviews %d' % movie_id)

    return sheet


def get_movie_details_cell(url, title, sheet: xlwt.Worksheet, movie_id: int, is_first=False):

    print('url = ', url)

    res = requests.get(url)

    assert res.status_code == 200

    res.encoding = 'utf-8'

    soup = BeautifulSoup(res.text, 'lxml')

    genre = soup.find('span', {'class':'ipc-chip__text'}, {'role': 'presentation'}).text

    item = soup.select('.ipc-metadata-list-item__content-container')

    director_link = url + item[0].find('a')['href']
    director = item[0].text

    writers = item[1].text

    stars = item[2].text

    storyline = soup.find('div', class_='ipc-html-content ipc-html-content--base').text

    row = [movie_id, title, genre, director, writers, stars, storyline]
    for i in range(len(row)):
        sheet.write(int(movie_id + 1), i, row[i])

    return sheet

def get_movie_details(url, title, sheet: xlwt.Worksheet, movie_id: int):

    sheet = get_movie_details_cell(url, title, sheet, movie_id, is_first=True)

    print('Finish movie details %d' % movie_id)

    return sheet

def write_sheet(file: xlwt.Workbook, sheet_name: str, row: List[str]):

    sheet: xlwt.Worksheet = file.add_sheet(sheet_name, cell_overwrite_ok=True)
    for i in range(len(row)):
        sheet.write(0, i, row[i])

    return sheet


def crawl_from_rank(url, sheet: List[xlwt.Worksheet]):

    print(url)

    res = requests.get(url)
    assert res.status_code == 200
    res.encoding = 'utf-8'
    soup = BeautifulSoup(res.text, 'lxml')

    sheet1, sheet2 = sheet

    for movie_id, item in enumerate(soup.select('.titleColumn')):

        title = item.find('a').text
        velocity = item.select('.velocity')[0]

        hot = re.findall(r"\d+", velocity.text)
        if len(hot) == 2:
            hot_rank, hot_trend = hot
            hot_trend = hot_trend if velocity.find_all('span', {'class':'global-sprite titlemeter up'}) != [] else "-" + hot_trend
        else:
            hot_rank = hot[0]
            hot_trend = 0

        movie_link = item.find('a')['href']
        movie_url = "https://www.imdb.com/" + movie_link

        sheet1 = get_movie_details(movie_url, title, sheet1, movie_id)

        movie_reviews_url = movie_url +  "reviews?ref_=tt_ov_rt"

        sheet2 = get_movie_reviews(movie_reviews_url, sheet2, movie_id)

        if movie_id + 1 >= MAX_MOVIE_CNT: break


def crawl_from_menu(url, sheet: xlwt.Worksheet):

    print('url = ', url)

    headers = {
        'Referer': 'https://www.imdb.com/',
        'sec-ch-ua': '"Google Chrome";v="95", "Chromium";v="95", ";Not A Brand";v="99"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': "Windows",
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.54 Safari/537.36',
        'x-amzn-sessionid': '131-3503887-4026530',
        'x-imdb-client-name': 'imdb-web-next',
        'x-imdb-user-country': 'CN',
        'x-imdb-user-language': 'zh-CN'
    }

    payload = {
        "query": "query FanFavorites($first: Int!, $after: ID) {\n  fanPicksTitles(first: $first, after: $after) {\n    edges {\n      node {\n        ...BaseTitleCard\n        ...TitleCardTrailer\n        ...TitleWatchOption\n        ...PersonalizedTitleCardUserRating\n        __typename\n      }\n      __typename\n    }\n    __typename\n  }\n}\n\nfragment BaseTitleCard on Title {\n  id\n  titleText {\n    text\n    __typename\n  }\n  titleType {\n    id\n    __typename\n  }\n  originalTitleText {\n    text\n    __typename\n  }\n  primaryImage {\n    id\n    width\n    height\n    url\n    __typename\n  }\n  releaseYear {\n    year\n    endYear\n    __typename\n  }\n  ratingsSummary {\n    aggregateRating\n    voteCount\n    __typename\n  }\n  runtime {\n    seconds\n    __typename\n  }\n  certificate {\n    rating\n    __typename\n  }\n  canRate {\n    isRatable\n    __typename\n  }\n  canHaveEpisodes\n}\n\nfragment TitleCardTrailer on Title {\n  latestTrailer {\n    id\n    __typename\n  }\n}\n\nfragment PersonalizedTitleCardUserRating on Title {\n  userRating {\n    value\n    __typename\n  }\n}\n\nfragment TitleWatchOption on Title {\n  primaryWatchOption {\n    additionalWatchOptionsCount\n    __typename\n  }\n}\n",
        "operationName": "FanFavorites",
        "variables": {"first":48}
    }

    res = requests.post(url, json=payload, headers=headers)

    assert res.status_code == 200

    data = json.loads(res.text)
    base_json = data['data']['fanPicksTitles']['edges']
    num_movies = len(base_json)
    for i in range(num_movies):
        title = base_json[i]['node']['titleText']['text']
        movie_id = base_json[i]['node']['id']

        key = '?ref_=watch_fanfav_tt_t_'
        url = 'https://www.imdb.com/'
        nextLink = url + 'title/' + movie_id + '/' + key + str(int(i)+1)
        # print('nextLink = ', nextLink)

        get_movie_details(nextLink, title, sheet, i)

def args_register():

    parser = argparse.ArgumentParser()
    parser.add_argument('--MAX_REVIEW_CNT', default=40, type=int, help="Maximum quantity of reviews from each movie.")
    parser.add_argument('--MAX_MOVIE_CNT', default=20, type=int, help='Maximum number of movies.')
    parser.add_argument('--CRAWL_FROM_RANK', action='store_false', help='Begining source for crawling.')
    parser.add_argument('--SAVE_PATH', default='IMDB_Reviews.xls', type=str, help='File saving path.')

    args = parser.parse_args()

    return args

f = xlwt.Workbook(encoding='utf-8')
row = ['Movie_ID', 'Title', 'Genre', 'Directors', 'Writers', 'Stars', 'StoryLine']
sheet1 = write_sheet(file=f, sheet_name='Movie Details', row=row)
row = ['Review_ID', 'Movie_ID', 'Title', 'Author', 'Date', 'Up Vote', 'Total Vote', 'Rating', 'Review']
sheet2 = write_sheet(file=f, sheet_name='Movie Reviews', row=row)

args = args_register()

MAX_REVIEW_CNT = args.MAX_REVIEW_CNT
MAX_MOVIE_CNT = args.MAX_MOVIE_CNT
CRAWL_FROM_RANK = args.CRAWL_FROM_RANK

if CRAWL_FROM_RANK:
    print('Crawling from rank...')
    crawl_from_rank(url='https://www.imdb.com/chart/moviemeter/?sort=ir,desc&mode=simple&page=1', sheet=[sheet1, sheet2])
else:
    print('Crawling from menu...')
    crawl_from_menu(url='https://api.graphql.imdb.com/', sheet=sheet1)

f.save(args.SAVE_PATH)
print('Crawling Finished')





