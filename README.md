# Author: KAM Wright
# Description: This code pull information on tweets related to screennames and outputs into an Excel file
# Version: 1
  
import openpyxl as xl
import datetime
import sys
import tweepy

consumer_key = '(INSERT YOUR KEY HERE)'
consumer_secret = '(INSERT YOUR KEY HERE)'
access_key = '(INSERT YOUR KEY HERE)'
access_secret = '(INSERT YOUR KEY HERE)'


def get_all_tweets(screen_name, end_date, limit=None):
    auth = tweepy.OAuthHandler(consumer_key, consumer_secret)
    auth.set_access_token(access_key, access_secret)
    api = tweepy.API(auth)
    	
    alltweets = []
    	
    try:
        new_tweets = api.user_timeline(screen_name = screen_name,count=200)
    except:
        print("Couldn't find tweets for user: {0}".format(screen_name))
        return None
    
    alltweets.extend(new_tweets)
    oldest = alltweets[-1].id - 1
    	
    while len(new_tweets) > 0:
        print "Getting tweets for {0}, {1} tweets dowloaded so far...".format(screen_name, len(alltweets))
        new_tweets = api.user_timeline(screen_name = screen_name,count=200,max_id=oldest)
        alltweets.extend(new_tweets)        
        oldest = alltweets[-1].id - 1
        
        if limit is not None:
            if len(alltweets) > limit:
                alltweets = alltweets[:limit]
                break

        if alltweets[-1].created_at <= (end_date - datetime.timedelta(days = 7)):
            break
            
    print "Finished getting {0} tweets from {1}".format(len(alltweets),screen_name)
     
    return filter_week(alltweets, end_date) 
         
def filter_week(alltweets, end_date):
    
    start_date = end_date - datetime.timedelta(days = 7)
    return  [tweet for tweet in alltweets if start_date <= tweet.created_at <= end_date]
         
def create_tweet_sheet(tweets, wb, user, headings):

    ws = wb.create_sheet(0)
    ws.title = user

    for row in range(0, len(tweets)+1):

        if row == 0:
            for col, heading in enumerate(headings):
                wsc = ws.cell(column = col+1, row=row+1)
                wsc.value = heading
        else:
            text = tweets[row-1].text
            retweet = tweets[row-1].retweeted or text[:2] == u'RT'
                
            wsc = ws.cell(column = 1, row=row+1)
            wsc.value = tweets[row-1].created_at
                
            wsc = ws.cell(column = 2, row=row+1)
            wsc.value = text
                
            wsc = ws.cell(column = 4, row=row+1)
            wsc.value = tweets[row-1].retweet_count
                
            wsc = ws.cell(column = 5, row=row+1)
            wsc.value = tweets[row-1].favorite_count
                
            wsc = ws.cell(column = 6, row=row+1)
            wsc.value = retweet
                
            wsc = ws.cell(column = 7, row=row+1)
            wsc.value = tweets[row-1].is_quote_status
                
            wsc = ws.cell(column = 8, row=row+1)
            wsc.value = tweets[row-1].author.followers_count
                
            wsc = ws.cell(column = 9, row=row+1)
            wsc.value = tweets[row-1].author.friends_count

def main(args):

    # First lets define a cut off date and user list
    cut_off_date = datetime.datetime(int(args[0]),int(args[1]),int(args[2]),
                                     int(args[3]),int(args[4]),int(args[5]))
    users = ['ConservativesIN','reformineurope','consforbritain','grassroots_out','labour4europe','Scientists4EU','labourleave','StrongerIn','LeaveEUOfficial','vote_leave']
    headings = ['Date','Tweet','Themes','Retweets','Favourites','Retweet?','Quoted?','Followers','Following','Retweets-retweeted','favourites-retweeted']
    
    # Now lets create an excel spreadsheet for the data  
    wb = xl.Workbook()

    # Loop over the users and get their tweets from 7 days before the cutoff
    for user in users:

        # Get all the tweets
        all_tweets = get_all_tweets(user, cut_off_date)
        if all_tweets is None: continue

        # Create the worksheet and give it a name and add some column heacdings
        create_tweet_sheet(all_tweets, wb, user, headings)

    # To Do: Conduct overall analysis and produce figures for results
    wb.save("C:\\Users\\kw0020\\TwitterData\\{0}.xlsx".format(cut_off_date.date()))

if __name__ == '__main__':
main(sys.argv[1:])
