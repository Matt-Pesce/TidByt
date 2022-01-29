load("render.star", "render")
load("http.star", "http")
load("cache.star", "cache")
load("time.star", "time")

# The daily step goal - this determines the coaching message you receive
STEP_GOAL = 10000

GOOGLEFIT_DATASET_URL = "https://www.googleapis.com/fitness/v1/users/me/dataset:aggregate"

#GOOGLEFIT_POSTREQUEST_BODY = {
#  "aggregateBy": [{
#    "dataSourceId":
#      "derived:com.google.step_count.delta:com.google.android.gms:estimated_steps"
#  }],
#  "bucketByTime": { "durationMillis": 86400000 },
#  "startTimeMillis": 1642914000000,
#  "endTimeMillis": 1642983862000
#}


GOOGLEFIT_OAUTH_TOKEN = "Bearer ya29.A0ARrdaM_v8NBQEs8sBbpOiOy806a395ehBrStNOuIYNjqRoFtMu_svvUaPJ6ji6yQesQ1Mr3SsdGNB2QufyXSZiwmFk-jrl0BolE9zeoFbqAmXJd9H-TMwELj39zrwUFaALKEAbbJ1tYFVw3PTziZtP3-8xmR"

GOOGLEFIT_POST_HEADERS = {"Content-type": "application/json",
  "Authorization": GOOGLEFIT_OAUTH_TOKEN
}


def main():

	# Current Time (Seconds)
	epoch_time = time.now().unix

	hour_time = time.now().hour
	min_time = time.now().minute
	second_time = time.now().second

	# Subtract out seconds to get the midnight Epoch time
	midnight_delta = hour_time*3600 + min_time*60 + second_time
	midnight_epoch = epoch_time - midnight_delta

	# Translate to milliseconds
	epoch_time_millis = epoch_time*1000
	midnight_epoch_millis = midnight_epoch*1000

	print("Time Now:\t %d ms" % epoch_time_millis)
	print("Midnight Today:\t %d ms" % midnight_epoch_millis)

	GOOGLEFIT_POSTREQUEST_BODY = {
  	  "aggregateBy": [{
    	    "dataSourceId":
      	      "derived:com.google.step_count.delta:com.google.android.gms:estimated_steps"
  	}],
  	"bucketByTime": { "durationMillis": 86400000 },
  	"startTimeMillis": midnight_epoch_millis,
  	"endTimeMillis": epoch_time_millis
	}

	rep = http.post(GOOGLEFIT_DATASET_URL,headers=GOOGLEFIT_POST_HEADERS,json_body=GOOGLEFIT_POSTREQUEST_BODY)
	if rep.status_code != 200:
		fail("Google FIT API request failed with status:", rep.json()["error"]["message"])

#	print(rep.json())
	step_count = rep.json()["bucket"][0]["dataset"][0]["point"][0]["value"][0]["intVal"]
	print("Step Count:", step_count)

	#Now select the appropriate "Coaching" message depending on Progress toward step count goal.  Count from 8am-midnight
	if hour_time > 8:
		progress_percent_expected = (hour_time - 8)/16
	else:
		progress_percent_expected = 0

	steps_expected = STEP_GOAL*progress_percent_expected
	steps_delta = step_count - steps_expected
	
	if steps_delta > 2000:
		coaching_msg = "Rock Star!"
		coaching_color = "#0f0"  	#Green
	elif steps_delta > 0:
		coaching_msg = "Very Good"
		coaching_color = "#fff"	#White
	elif steps_delta <= -2000:
		coaching_msg = "Slug!!"
		coaching_color = "#00f"	#blue
	else:
		coaching_msg = "Get Moving"
		coaching_color = "#f00"	#Red
	
	print("Expected Steps:", steps_expected)
	print(coaching_msg)			

# Old code attempting to use time parsing function (parsing function choking on whitespace...or ???
# Also, time parsing function always returns Zulu (EST +5:00) even if location parameter is used (is it misspelled?)
# putting -5:00 after the Z also causes a syntax error
#	epoch_time = time.now().unix
#	now = str(time.now())
#	date_portion = now.rsplit(" ")
#	print(now)
#
#	day_time = str(time.now().day)
#	year_time = str(time.now().year)
#	month_time = str(time.now().month)
#	
#
##	midnight_today = "-".join([day_time, month_time, year_time, "00:00:00"])  <--Cannot get Parser to like this (single digits, likely)
#	midnight_today = "T".join([date_portion[0], "00:00:00Z"])
#	midnight_epoch = time.parse_time(midnight_today, format="2022-12-22 01:02:03", location="US/Eastern").unix
#	midnight_epoch = time.parse_time(date_portion[0], location="US/Eastern").unix
##	metric = time.parse_time("2022-01-22T00:00:00Z").unix   <-hard code benchmark.  This works.
#	midnight_epoch = time.parse_time(midnight_today, location="US/Eastern").unix
	
    	
   	return render.Root(
		render.Column(
			main_align="space_evenly",
			cross_align="center",
			expanded=True,
			children=[
#				render.Marquee(
#					width=64,
#					child = render.Text("%s" % quote),
#				),
				render.Text("   Step Count   "),
				render.Text("%d" % step_count),
				render.Text(""),
				render.Text("%s" % coaching_msg, color=coaching_color),
#				render.Text(content="-%s" % author, color="#f00", font="CG-pixel-3x5-mono")
			],
		)
     	)	

