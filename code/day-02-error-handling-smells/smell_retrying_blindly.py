def fetch_exchange_rate(currency):
    while True:
        try:
            return exchange_api.get_rate(currency)
        except ConnectionError:
            continue

# no limit, no backoff - if the API is down, this
# hammers it as fast as the network allows, forever
