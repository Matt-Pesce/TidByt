load("render.star", "render")
load("http.star", "http")
load("cache.star", "cache")

ZENQUOTE_RANDOM_URL = "https://zenquotes.io/api/random"

def main():
	quote_cached = cache.get("quote_rate")
	author_cached = cache.get("author_rate")
	print("Quote Cached: %s" % quote_cached)
	if quote_cached != None:
		print("Fetching Cached Quote")
		quote = (quote_cached)
		author = (author_cached)
	else:
		print("Fetching from Zenquote")
		rep = http.get(ZENQUOTE_RANDOM_URL)
		print ("Rep: %s" % rep)
		if rep.status_code != 200:
			fail("Zenquote request failed with status %d", rep.status_code)
		quote = rep.json()[0]["q"]
		author = rep.json()[0]["a"]
		print ("%s" % quote)
		cache.set("quote_rate", quote, ttl_seconds=30)
		cache.set("author_rate", author, ttl_seconds=30)     #MF note, will there be a race condition - possibility of mixed quote/author?
    	
    	return render.Root(
		render.Column(
			main_align="space_evenly",
			children=[
				render.Marquee(
					width=64,
					child = render.Text("%s" % quote),
				),
				render.Text(content=""),
				render.Text(content="-%s" % author, color="#f00", font="CG-pixel-3x5-mono"),
			],
		)
      	)	
 
