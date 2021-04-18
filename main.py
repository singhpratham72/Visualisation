from bs4 import BeautifulSoup
import requests
import xlsxwriter

class Hotel:

    def __init__(self, name, link):
        self.name = name
        self.link = link
        self.overallRating = 0
        self.reviews = []
        self.ratings = []

    def addData(self, overall, reviewList, ratingList):
        self.overallRating = overall
        self.reviews = reviewList
        self.ratings = ratingList
        
# Return a list of hotels from the url provided  
def getHotels(url):
    html_text = requests.get(url).content
    soup = BeautifulSoup(html_text, 'lxml')

    # List of hotels
    hotels = []

    # For all hotels
    cards = soup.find_all('div', class_ = 'prw_rup prw_meta_hsx_responsive_listing ui_section listItem')
    # To store name and link of hotel in Hotel object
    for card in cards:
        if(card == cards[20]):
            break
        try:
            hotelObj = []
            hotelObj = card.select('h2', class_ = 'property_title prominent')
            # if-else as the website code was different for some cities
            if(len(hotelObj) == 0):
                hotelObj = card.find('div', class_ = 'listing_title')
                hotelName = hotelObj.text.strip()
                hotelLink = 'https://www.tripadvisor.in' + hotelObj.a['href'] + '#REVIEWS'
            else:
                hotelName = hotelObj[0].text.strip()
                hotelLink = 'https://www.tripadvisor.in' + hotelObj[0].a['href'] + '#REVIEWS'
            hotel = Hotel(name = hotelName, link = hotelLink)
        except Exception as e:
            print(e)
            continue
        hotels.append(hotel)

    return hotels

# Return a tuple of overallRating, reviews[], ratings[] of a hotel
def getHotelData(link):
    reviewList = []
    ratingList = []
    overall = 0
    reviewPage = ''

    # Loop to keep shifting to the next page of reviews
    for r_num in range (0, 100, 5):
        
        if (r_num != 0):
            reviewPage = '-or' + f'{r_num}'

        li = link.split('Reviews')
        url = li[0] + 'Reviews' + reviewPage + li[1]

        html_text = requests.get(url).content
        soup = BeautifulSoup(html_text, 'lxml')

        # Overall Rating
        about = soup.find('span', class_ = '_3cjYfwwQ')
        overall = about.text

        # Reviews
        reviews = soup.find_all('div', class_ = '_2wrUUKlw _3hFEdNs8')

        for review in reviews:

            # Review Rating
            ratingObj = review.find('div', class_ = 'nf9vGX55')
            rating = 0

            str = ratingObj.find('span', class_ = 'ui_bubble_rating bubble_50')
            if (str != None):
                rating = 5
            str = ratingObj.find('span', class_ = 'ui_bubble_rating bubble_40')
            if (str != None):
                rating = 4
            str = ratingObj.find('span', class_ = 'ui_bubble_rating bubble_30')
            if (str != None):
                rating = 3
            str = ratingObj.find('span', class_ = 'ui_bubble_rating bubble_20')
            if (str != None):
                rating = 2
            str = ratingObj.find('span', class_ = 'ui_bubble_rating bubble_10')
            if (str != None):
                rating = 1

            # Review Comment
            commentObj = review.find('q', class_ = 'IRsGHoPm')
            parts = commentObj.find_all('span')
            comment = ''
            # To add hidden parts of reviews
            for part in parts:
                comment = comment + part.text

            reviewList.append(comment)
            ratingList.append(rating)

            # print (f'Rating = {rating} bubbles\nReview: {comment}\n-----------------')
    return (overall, reviewList, ratingList)

def main():

    # To fetch hotels in Bangkok
    workbook = xlsxwriter.Workbook('bangkokHotels.xlsx')
    worksheet = workbook.add_worksheet()

    row = 1
    col = 0

    worksheet.write(0, 0, 'Hotel Name')
    worksheet.write(0, 1, 'Hotel Link')
    worksheet.write(0, 2, 'Overall')
    worksheet.write(0, 3, 'Rating')
    worksheet.write(0, 4, 'Review')

    kualaLumpurHotels = getHotels(url = 'https://www.tripadvisor.in/Hotels-g293916-Bangkok-Hotels.html')
    for hotel in kualaLumpurHotels:
        hotelData = getHotelData(hotel.link)
        hotel.addData(overall = hotelData[0], reviewList = hotelData[1], ratingList = hotelData[2])
        for i in range (len(hotel.reviews)):
            worksheet.write(row, col, hotel.name)
            worksheet.write(row, col + 1, hotel.link)
            worksheet.write(row, col + 2, hotel.overallRating)
            worksheet.write(row, col + 3, hotel.ratings[i])
            worksheet.write(row, col + 4, hotel.reviews[i])
            row = row + 1
        
    workbook.close()

    # To fetch hotels in Kuala Lumpur
    workbook = xlsxwriter.Workbook('kualaLumpurHotels.xlsx')
    worksheet = workbook.add_worksheet()

    row = 1
    col = 0

    worksheet.write(0, 0, 'Hotel Name')
    worksheet.write(0, 1, 'Hotel Link')
    worksheet.write(0, 2, 'Overall')
    worksheet.write(0, 3, 'Rating')
    worksheet.write(0, 4, 'Review')

    kualaLumpurHotels = getHotels(url = 'https://www.tripadvisor.in/Hotels-g298570-Kuala_Lumpur_Wilayah_Persekutuan-Hotels.html')
    for hotel in kualaLumpurHotels:
        hotelData = getHotelData(hotel.link)
        hotel.addData(overall = hotelData[0], reviewList = hotelData[1], ratingList = hotelData[2])
        for i in range (len(hotel.reviews)):
            worksheet.write(row, col, hotel.name)
            worksheet.write(row, col + 1, hotel.link)
            worksheet.write(row, col + 2, hotel.overallRating)
            worksheet.write(row, col + 3, hotel.ratings[i])
            worksheet.write(row, col + 4, hotel.reviews[i])
            row = row + 1
        
    workbook.close()

    # To fetch hotels in Singapore
    workbook = xlsxwriter.Workbook('singaporeHotels.xlsx')
    worksheet = workbook.add_worksheet()

    row = 1
    col = 0

    worksheet.write(0, 0, 'Hotel Name')
    worksheet.write(0, 1, 'Hotel Link')
    worksheet.write(0, 2, 'Overall')
    worksheet.write(0, 3, 'Rating')
    worksheet.write(0, 4, 'Review')

    singaporeHotels = getHotels(url = 'https://www.tripadvisor.in/Hotels-g294265-Singapore-Hotels.html')
    for hotel in singaporeHotels:
        hotelData = getHotelData(hotel.link)
        hotel.addData(overall = hotelData[0], reviewList = hotelData[1], ratingList = hotelData[2])
        for i in range (len(hotel.reviews)):
            worksheet.write(row, col, hotel.name)
            worksheet.write(row, col + 1, hotel.link)
            worksheet.write(row, col + 2, hotel.overallRating)
            worksheet.write(row, col + 3, hotel.ratings[i])
            worksheet.write(row, col + 4, hotel.reviews[i])
            row = row + 1
        
    workbook.close()


main()