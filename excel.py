
import pandas as pd
import json
import os

# Read the JSON file
def export_to_excel():
    with open('results.json') as file:
        data = json.load(file)

    # Create a list to store the values of the specified key
    title = []
    length = []
    width = []
    total_height = []
    age_group = []
    maximum_weight_of_user = []
    safety_zone = []
    number_of_users = []
    free_fall_height = []
    in_accordance_with_norm_EN = []
    weight_of_the_heaviest_part = []
    dimensions_of_the_largest_part = []
    availability_of_spare_parts = []
    installation_time = []
    image_url = []

    # Iterate over each item in the JSON data
    for item in data:
        # Get the value of the specified key
        title.append(item.get('title'))
        length.append(item.get('Length'))
        width.append(item.get('Width'))
        total_height.append(item.get('Total height'))
        age_group.append(item.get('Age group'))
        number_of_users.append(item.get('Number of users'))
        maximum_weight_of_user.append(item.get('Maximum weight of user'))
        safety_zone.append(item.get('Safety zone'))
        free_fall_height.append(item.get('Free fall height'))
        in_accordance_with_norm_EN.append(item.get('In accordance with norm EN'))
        weight_of_the_heaviest_part.append(item.get('Weight of the heaviest part'))
        dimensions_of_the_largest_part.append(item.get('Dimensions of the largest part'))
        availability_of_spare_parts.append(item.get('Availability of spare parts'))
        installation_time.append(item.get('Installation time'))
        image_url.append(item.get('url'))

    # Create a DataFrame from the values of the specified keys
    df = pd.DataFrame({
        "title": title,
        "Length": length,
        "Width": width,
        "Total height": total_height,
        "Age group": age_group,
        "Maximum weight of user": maximum_weight_of_user,
        "Number of users" :  number_of_users,
        "Safety zone": safety_zone,
        "Free fall height": free_fall_height,
        "In accordance with norm EN": in_accordance_with_norm_EN,
        "Weight of the heaviest part": weight_of_the_heaviest_part,
        "Dimensions of the largest part": dimensions_of_the_largest_part,
        "Availability of spare parts": availability_of_spare_parts,
        "Installation time": installation_time,
        "Image url" : image_url
    })

    filename = os.path.join(".", "vinci.xlsx")
    df.to_excel(filename, index=False)

    print("Exported successfully to:", filename)

export_to_excel()