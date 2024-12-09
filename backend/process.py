import json

def find_deferred_revenue(data):
    results = []

    def search(item):
        # If item is a dictionary, search its values
        if isinstance(item, dict):
            # Check for the specific pattern indicating "Deferred Revenue"
            for key, value in item.items():
                if key == 'ColData' and isinstance(value, list):
                    # Check if "Deferred Revenue" is in the ColData list
                    for col in value:
                        if col.get('value') == 'Deferred Revenue':
                            # Find the next dictionary in the list and get the associated amount
                            idx = value.index(col)
                            if idx + 1 < len(value):
                                amount = value[idx + 1].get('value')
                                results.append(amount)
                else:
                    search(value)
        # If item is a list, search each element
        elif isinstance(item, list):
            for element in item:
                search(element)

    # Start searching from the root of the JSON object
    search(data)

    return results

# Load JSON data from a file
with open('def-rev-report.json', 'r') as file:
    data = json.load(file)

# Find all occurrences of "Deferred Revenue" and their associated amounts
deferred_revenue_values = find_deferred_revenue(data)

# Print the results
print("Deferred Revenue Amounts:")
for amount in deferred_revenue_values:
    print(amount)