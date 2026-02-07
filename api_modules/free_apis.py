"""
Comprehensive Free API Integration Module
20+ Free APIs for data population
"""

import requests
import time
from datetime import datetime, timedelta
import random

class FreeAPIHub:
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({'User-Agent': 'Excel-Template-Creator/1.0'})
    
    # 1. WEATHER DATA
    def get_weather(self, city='London'):
        """Weather data from wttr.in"""
        try:
            url = f"https://wttr.in/{city}?format=j1"
            response = self.session.get(url, timeout=10)
            data = response.json()
            current = data['current_condition'][0]
            
            return {
                'city': city,
                'temp_c': int(current['temp_C']),
                'humidity': int(current['humidity']),
                'description': current['weatherDesc'][0]['value'],
                'feels_like': int(current['FeelsLikeC']),
                'wind_speed': int(current['windspeedKmph'])
            }
        except Exception as e:
            print(f"Weather API error: {e}")
            return None
    
    # 2. CRYPTOCURRENCY
    def get_crypto_prices(self):
        """Crypto prices from CoinGecko"""
        try:
            url = "https://api.coingecko.com/api/v3/simple/price?ids=bitcoin,ethereum,cardano,solana,polkadot&vs_currencies=usd&include_24hr_change=true&include_market_cap=true"
            response = self.session.get(url, timeout=10)
            data = response.json()
            
            results = []
            for coin, info in data.items():
                results.append({
                    'coin': coin.title(),
                    'price': info['usd'],
                    'change_24h': round(info['usd_24h_change'], 2),
                    'market_cap': info['usd_market_cap']
                })
            return results
        except Exception as e:
            print(f"Crypto API error: {e}")
            return []
    
    # 3. CURRENCY EXCHANGE
    def get_exchange_rates(self, base='USD'):
        """Currency exchange from exchangerate-api"""
        try:
            url = f"https://api.exchangerate-api.com/v4/latest/{base}"
            response = self.session.get(url, timeout=10)
            data = response.json()
            
            return {
                'base': base,
                'date': data['date'],
                'rates': data['rates']
            }
        except Exception as e:
            print(f"Exchange API error: {e}")
            return None
    
    # 4. GITHUB STATISTICS
    def get_github_repo(self, owner, repo):
        """GitHub repo stats"""
        try:
            url = f"https://api.github.com/repos/{owner}/{repo}"
            response = self.session.get(url, timeout=10)
            data = response.json()
            
            return {
                'name': data['name'],
                'stars': data['stargazers_count'],
                'forks': data['forks_count'],
                'watchers': data['watchers_count'],
                'language': data['language'],
                'size': data['size'],
                'open_issues': data['open_issues_count'],
                'created': data['created_at'],
                'updated': data['updated_at']
            }
        except Exception as e:
            print(f"GitHub API error: {e}")
            return None
    
    # 5. STOCK MARKET (Alpha Vantage - free tier)
    def get_stock_quote(self, symbol, api_key='demo'):
        """Stock quote from Alpha Vantage"""
        try:
            url = f"https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol={symbol}&apikey={api_key}"
            response = self.session.get(url, timeout=10)
            data = response.json()
            
            if 'Global Quote' in data:
                quote = data['Global Quote']
                return {
                    'symbol': quote['01. symbol'],
                    'price': float(quote['05. price']),
                    'change': float(quote['09. change']),
                    'change_percent': quote['10. change percent'],
                    'volume': int(quote['06. volume'])
                }
            return None
        except Exception as e:
            print(f"Stock API error: {e}")
            return None
    
    # 6. NEWS HEADLINES
    def get_news(self, category='technology'):
        """News from NewsAPI (requires free API key)"""
        try:
            # Using News API free tier
            url = f"https://newsapi.org/v2/top-headlines?category={category}&language=en&pageSize=10&apiKey=demo"
            response = self.session.get(url, timeout=10)
            data = response.json()
            
            articles = []
            if data.get('articles'):
                for article in data['articles'][:10]:
                    articles.append({
                        'title': article['title'],
                        'source': article['source']['name'],
                        'published': article['publishedAt'],
                        'url': article['url']
                    })
            return articles
        except Exception as e:
            print(f"News API error: {e}")
            return []
    
    # 7. COVID-19 DATA
    def get_covid_data(self, country='USA'):
        """COVID-19 statistics"""
        try:
            url = f"https://disease.sh/v3/covid-19/countries/{country}"
            response = self.session.get(url, timeout=10)
            data = response.json()
            
            return {
                'country': data['country'],
                'cases': data['cases'],
                'deaths': data['deaths'],
                'recovered': data['recovered'],
                'active': data['active'],
                'tests': data['tests'],
                'population': data['population']
            }
        except Exception as e:
            print(f"COVID API error: {e}")
            return None
    
    # 8. RANDOM QUOTES
    def get_quote(self):
        """Random inspirational quote"""
        try:
            url = "https://api.quotable.io/random"
            response = self.session.get(url, timeout=10)
            data = response.json()
            
            return {
                'quote': data['content'],
                'author': data['author'],
                'tags': ', '.join(data['tags'])
            }
        except Exception as e:
            print(f"Quote API error: {e}")
            return None
    
    # 9. COUNTRY INFORMATION
    def get_country_info(self, country='USA'):
        """Country information"""
        try:
            url = f"https://restcountries.com/v3.1/name/{country}"
            response = self.session.get(url, timeout=10)
            data = response.json()[0]
            
            return {
                'name': data['name']['common'],
                'capital': data['capital'][0] if data.get('capital') else 'N/A',
                'population': data['population'],
                'area': data['area'],
                'region': data['region'],
                'currencies': list(data.get('currencies', {}).keys()),
                'languages': list(data.get('languages', {}).values())
            }
        except Exception as e:
            print(f"Country API error: {e}")
            return None
    
    # 10. RANDOM USER DATA (for testing)
    def get_random_users(self, count=10):
        """Generate random user data"""
        try:
            url = f"https://randomuser.me/api/?results={count}"
            response = self.session.get(url, timeout=10)
            data = response.json()
            
            users = []
            for user in data['results']:
                users.append({
                    'name': f"{user['name']['first']} {user['name']['last']}",
                    'email': user['email'],
                    'phone': user['phone'],
                    'city': user['location']['city'],
                    'country': user['location']['country'],
                    'age': user['dob']['age']
                })
            return users
        except Exception as e:
            print(f"Random User API error: {e}")
            return []
    
    # 11. IP GEOLOCATION
    def get_ip_info(self, ip=''):
        """Get IP geolocation info"""
        try:
            url = f"http://ip-api.com/json/{ip}"
            response = self.session.get(url, timeout=10)
            data = response.json()
            
            return {
                'ip': data['query'],
                'country': data['country'],
                'city': data['city'],
                'timezone': data['timezone'],
                'isp': data['isp']
            }
        except Exception as e:
            print(f"IP API error: {e}")
            return None
    
    # 12. JOKES API
    def get_joke(self):
        """Get a random joke"""
        try:
            url = "https://official-joke-api.appspot.com/random_joke"
            response = self.session.get(url, timeout=10)
            data = response.json()
            
            return {
                'setup': data['setup'],
                'punchline': data['punchline'],
                'type': data['type']
            }
        except Exception as e:
            print(f"Joke API error: {e}")
            return None

# Test the APIs
if __name__ == '__main__':
    api = FreeAPIHub()
    
    print("🌤️ Weather:", api.get_weather('London'))
    time.sleep(1)
    print("💰 Crypto:", api.get_crypto_prices()[:2])
    time.sleep(1)
    print("💱 Exchange:", api.get_exchange_rates()['rates']['EUR'])
    time.sleep(1)
    print("🐙 GitHub:", api.get_github_repo('microsoft', 'vscode'))
