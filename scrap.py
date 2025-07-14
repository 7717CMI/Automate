import requests
import json
import time

class SimpleTrendsExtractor:
    def __init__(self, api_key):
        self.api_key = api_key
        self.base_url = "https://api.scrapingdog.com/google_trends"
    
    def get_values(self, keyword):
        """Get exact values for last 7 days only"""
        print(f"🔍 Getting 7-day values for: '{keyword}'")
        
        params = {
            "api_key": self.api_key,
            "query": keyword,
            "geo": "",           # Worldwide
            "tz": "330",         # Indian timezone (UTC+5:30)
            "date": "now 7-d",   # Last 7 days
            "data_type": "TIMESERIES"
        }
        
        try:
            response = requests.get(self.base_url, params=params)
            
            if response.status_code == 200:
                data = response.json()
                values = self.extract_values(data)
                
                if values:
                    print(f"   ✅ Found 7-day values: {values}")
                    return values
                else:
                    print(f"   ❌ No 7-day values found")
                    return []
            else:
                print(f"   ❌ API Error: {response.status_code}")
                try:
                    error_data = response.json()
                    print(f"   Error details: {error_data}")
                except:
                    print(f"   Error text: {response.text}")
                return []
        
        except Exception as e:
            print(f"   ❌ Exception: {e}")
            return []
    
    def extract_values(self, data):
        """Extract values using only the standard timeline method"""
        values = []
        
        try:
            if 'interest_over_time' in data:
                timeline_data = data['interest_over_time'].get('timeline_data', [])
                
                for entry in timeline_data:
                    if isinstance(entry, dict) and 'values' in entry:
                        for val_item in entry['values']:
                            if isinstance(val_item, dict) and 'value' in val_item:
                                try:
                                    val = int(val_item['value'])
                                    if 0 <= val <= 100:
                                        values.append(val)
                                except (ValueError, TypeError):
                                    pass
        except Exception:
            pass
        
        return values
    
    def process_keywords(self, keywords):
        """Process keywords - get 7-day values and filter if 2+ values > 50"""
        print(f"🚀 Processing {len(keywords)} keywords with 7-day filter")
        print("=" * 60)
        
        results = {}
        accepted = []
        rejected = []
        
        for i, keyword in enumerate(keywords, 1):
            print(f"\n[{i}/{len(keywords)}] {keyword}")
            
            # Get 7-day values
            values = self.get_values(keyword)
            
            if values:
                # Check filter: count values > 50
                count_above_50 = sum(1 for val in values if val > 50)
                
                if count_above_50 >= 2:
                    print(f"   ✅ ACCEPTED: {count_above_50} values > 50 in {values}")
                    results[keyword] = values
                    accepted.append(keyword)
                else:
                    print(f"   ❌ REJECTED: Only {count_above_50} values > 50 in {values}")
                    rejected.append(keyword)
            else:
                print(f"   ❌ REJECTED: No data")
                rejected.append(keyword)
            
            # Rate limiting
            if i < len(keywords):
                time.sleep(1)
        
        print(f"\n📊 FILTER SUMMARY")
        print("=" * 60)
        print(f"✅ Accepted: {len(accepted)} keywords")
        print(f"❌ Rejected: {len(rejected)} keywords")
        
        return results
    
    def display_results(self, results):
        """Display only accepted keywords"""
        print(f"\n📊 ACCEPTED KEYWORDS (7-Day Values)")
        print("=" * 60)
        
        if results:
            for keyword, values in results.items():
                print(f"✅ {keyword}: {values}")
        else:
            print("❌ No keywords passed the filter")
        
        return results


def main():
    api_key = "687286d843d158b2e5b064a9"
    extractor = SimpleTrendsExtractor(api_key)
    
    keywords = [
        "Public Cloud",
        "machine learning"
    ]
    
    # Process keywords
    results = extractor.process_keywords(keywords)
    
    # Display results
    final_results = extractor.display_results(results)
    
    # Return the values
    return final_results


if __name__ == "__main__":
    main()
