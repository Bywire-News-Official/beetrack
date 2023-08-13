import csv
import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit as st

if "st" in globals():
    RUN_IN_STREAMLIT = True
else:
    RUN_IN_STREAMLIT = False

def create_csv(filename):
    with open(filename, "w", newline='') as csvfile:
        writer = csv.writer(csvfile)
        header = ["Swarm ID", "Swarm Week", "Swarm Number", "Swarm URL", "Tweet Content", "Tweet Image File Name",
                  "Views", "Retweets", "Quotes", "Likes", "Bookmarks", "Comments",
                  "Ending Views", "Ending Retweets", "Ending Quotes", "Ending Likes", "Ending Bookmarks"]
        writer.writerow(header)

def add_swarm(filename, swarm_data):
    sw_id = f"SW-{swarm_data['sw_week']}-{swarm_data['sw_number']}"
    with open(filename, "a", newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow([sw_id, swarm_data['sw_week'], swarm_data['sw_number'], swarm_data['sw_url'], swarm_data['tweet_content'],
                         swarm_data['tweet_image_name'], swarm_data['views'], swarm_data['retweets'], swarm_data['quotes'],
                         swarm_data['likes'], swarm_data['bookmarks'], swarm_data['comments'], swarm_data['ending_views'],
                         swarm_data['ending_retweets'], swarm_data['ending_quotes'], swarm_data['ending_likes'], swarm_data['ending_bookmarks']])
        return sw_id

def edit_swarm(filename, swarm_id, column_to_edit, new_value):
    swarm = None
    with open(filename, newline='') as csvfile:
        reader = csv.reader(csvfile)
        header = next(reader)
        for row in reader:
            if row[0] == swarm_id:
                swarm = row
                break

    if not swarm:
        return f"Swarm ID {swarm_id} not found."

    swarm[column_to_edit] = new_value

    with open(filename, newline='') as csvfile:
        reader = list(csv.reader(csvfile))
        index_to_update = next(i for i, row in enumerate(reader) if row[0] == swarm_id)
        reader[index_to_update] = swarm

    with open(filename, "w", newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(reader)

    return f"Swarm {swarm_id} updated."

def list_swarms_and_comments(filename):
    with open(filename, newline='') as csvfile:
        reader = csv.reader(csvfile)
        next(reader)  # skip header row
        for row in reader:
            comments = row[11].split("\\n")
            first_comment = comments[0]
            st.write(f"{row[0]}: {first_comment}")
            if len(comments) > 1:
                if st.button("More", key=row[0]):
                    remaining_comments = "\\n".join(comments[1:])
                    st.write(remaining_comments)

if RUN_IN_STREAMLIT:
    def load_data(filename):
        try:
            df = pd.read_csv(filename, error_bad_lines=False)
            # Convert columns to appropriate data types
            numerical_cols = ["Views", "Retweets", "Quotes", "Likes", "Bookmarks",
                              "Ending Views", "Ending Retweets", "Ending Quotes", "Ending Likes", "Ending Bookmarks"]
            for col in numerical_cols:
                df[col] = pd.to_numeric(df[col], errors='coerce')
            return df
        except FileNotFoundError:
            return None

    def show_total_swarms(df):
        if df is not None:
            total_swarms = len(df)
            st.header(f"Total Number of Swarms: {total_swarms}")

    def show_total_comments(df):
        if df is not None:
            comments_list = df['Comments'].str.split('\\n').explode().dropna().tolist()
            st.header(f"Total Number of Comments: {len(comments_list)}")

    def plot_swarms_per_week(df):
        if df is not None and not df.empty:
            swarms_per_week = (
                df.groupby("Swarm Week")[["Swarm ID"]]
                .count()
                .reset_index()
                .rename(columns={"Swarm ID": "Count"})
            )

            if not swarms_per_week["Count"].isnull().all():
                f, ax = plt.subplots(figsize=(10, 6))
                sns.barplot(x="Swarm Week", y="Count", data=swarms_per_week, ax=ax)
                plt.title("Number of Swarms per Week")
                st.pyplot(f)
            else:
                st.write("No swarms data available for the given week.")
        else:
            st.write("No data available.")

    def plot_difference_per_week(df):
        if df is not None and not df.empty:
            difference_per_week = df.copy()
            difference_per_week['Retweets Difference'] = difference_per_week['Ending Retweets'] - difference_per_week['Retweets']
            difference_per_week['Quotes Difference'] = difference_per_week['Ending Quotes'] - difference_per_week['Quotes']
            difference_per_week['Likes Difference'] = difference_per_week['Ending Likes'] - difference_per_week['Likes']
            difference_per_week['Bookmarks Difference'] = difference_per_week['Ending Bookmarks'] - difference_per_week['Bookmarks']
            difference_per_week = difference_per_week.groupby('Swarm Week').sum().reset_index()
            difference_per_week = difference_per_week.melt(id_vars=['Swarm Week'], value_vars=['Retweets Difference', 'Quotes Difference', 'Likes Difference', 'Bookmarks Difference'])

            if not difference_per_week['value'].isnull().all():
                f, ax = plt.subplots(figsize=(10, 6))
                sns.barplot(x="Swarm Week", y="value", hue="variable", data=difference_per_week, ax=ax)
                plt.title("Engagement Metrics per Week (Before and After)")
                st.pyplot(f)
            else:
                st.write("No difference data available for the given week.")
        else:
            st.write("No data available.")

    def plot_views_difference_per_week(df):
        if df is not None and not df.empty:
            difference_per_week = df.copy()
            difference_per_week['Views Difference'] = difference_per_week['Ending Views'] - difference_per_week['Views']
            difference_per_week = difference_per_week.groupby('Swarm Week')[['Views', 'Ending Views']].sum().reset_index()

            if not difference_per_week['Views'].isnull().all() and not difference_per_week['Ending Views'].isnull().all():
                f, ax = plt.subplots(figsize=(10, 6))
                sns.barplot(x="Swarm Week", y="Views", data=difference_per_week, 
                            label="Beginning Views", color="skyblue", ax=ax)
                sns.barplot(x="Swarm Week", y="Ending Views", data=difference_per_week, 
                            label="Ending Views", color="royalblue", ax=ax)
                ax.legend()
                plt.title("Difference between Starting and Ending Views per Week")
                st.pyplot(f)
            else:
                                st.write("No views data available for the given week.")
        else:
            st.write("No data available.")

    st.title("Dashboard and Swarms CSV Management")

    filename = st.text_input("Enter the name of the CSV file or create a new one:")

    if os.path.exists(filename):
        st.write(f"{filename} loaded successfully!")

        data = load_data(filename)
        show_total_swarms(data)
        show_total_comments(data)
        plot_swarms_per_week(data)
        plot_views_difference_per_week(data)
        plot_difference_per_week(data)
        st.write("")
        
        st.write(f"{len(data)} swarms in {filename}")

    st.write("Add a new swarm:")
    swarm_data = {}
    
    def reset_form_fields():
        for key in swarm_data.keys():
            swarm_data[key] = None

    swarm_data['sw_week'] = st.text_input("Enter Swarm Week number:")
    swarm_data['sw_number'] = st.text_input("Enter Swarm Number:")
    swarm_data['sw_url'] = st.text_input("Enter Swarm URL:")
    swarm_data['tweet_content'] = st.text_input("Enter Tweet Content:")
    swarm_data['tweet_image_name'] = st.text_input("Enter Tweet Image File Name:")
    swarm_data['views'] = st.text_input("Enter Views:")
    swarm_data['retweets'] = st.text_input("Enter Retweets:")
    swarm_data['quotes'] = st.text_input("Enter Quotes:")
    swarm_data['likes'] = st.text_input("Enter Likes:")
    swarm_data['bookmarks'] = st.text_input("Enter Bookmarks:")
    swarm_data['comments'] = st.text_area("Enter Comments (separated by new lines):")
    swarm_data['ending_views'] = st.text_input("Enter Ending Views:")
    swarm_data['ending_retweets'] = st.text_input("Enter Ending Retweets:")
    swarm_data['ending_quotes'] = st.text_input("Enter Ending Quotes:")
    swarm_data['ending_likes'] = st.text_input("Enter Ending Likes:")
    swarm_data['ending_bookmarks'] = st.text_input("Enter Ending Bookmarks:")

    if st.button("Add Swarm"):
        comments_list = swarm_data['comments'].split('\\n')
        joined_comments = '\\n'.join(comments_list)
        swarm_data['comments'] = joined_comments
        
        current_swarm_id = add_swarm(filename, swarm_data)
        st.success(f"Swarm {current_swarm_id} added.")
        
        reset_form_fields()

    st.write("Edit a swarm:")
    swarm_id = st.text_input("Enter the Swarm ID to edit:", key='swarm_id')
    
    editable_columns = ["Swarm ID", "Swarm Week", "Swarm Number", "Swarm URL", "Tweet Content", "Tweet Image File Name",
                  "Views", "Retweets", "Quotes", "Likes", "Bookmarks", "Comments",
                  "Ending Views", "Ending Retweets", "Ending Quotes", "Ending Likes", "Ending Bookmarks"]
    selected_column_name = st.selectbox("Choose the column to edit:", editable_columns, key='selected_column_name')
    selected_column_index = editable_columns.index(selected_column_name)
    
    new_value = st.text_input("Enter the new value:", key='new_value')

    if st.button("Update Swarm"):
        result = edit_swarm(filename, swarm_id, selected_column_index, new_value)
        st.write(result)
    
    if filename and os.path.exists(filename):
        st.write("Current Swarms and Comments:")
        list_swarms_and_comments(filename)
    else:
        if filename:
            st.write(f"Creating a new CSV {filename}.")
            create_csv(filename)
else:
    print("This script is designed to be run in Streamlit. Please run it with `streamlit run script.py`.")

