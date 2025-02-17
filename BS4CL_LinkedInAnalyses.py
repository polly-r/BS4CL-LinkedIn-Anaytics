import streamlit as st
import pandas as pd
import altair as alt

st.set_page_config(layout="wide")
st.title("BS4CL LinkedIn Page 2025")

# =============================================================================
# Function to load Excel sheets from a file path
# =============================================================================
def load_excel_sheets(file_path):
    try:
        xls = pd.ExcelFile(file_path)
        # Load each sheet with no header; we’ll process headers later.
        return {sheet: xls.parse(sheet, header=None) for sheet in xls.sheet_names}
    except Exception as e:
        st.error(f"Error loading {file_path}: {e}")
        return {}

# =============================================================================
# Define file paths (update these paths as needed)
# =============================================================================
competitor_file = "Business Schools 4 Climate Leadership Africa_competitor_analytics_1739343948666.xlsx"
visitors_file   = "business-schools-4-climate-leadership-africa_visitors_1739343988681.xlsx"
followers_file  = "business-schools-4-climate-leadership-africa_followers_1739343974503.xlsx"

# Can improve via more metrics
# Metrics such as likes, comments, and shares for each post or update. 
# # Add a table or chart ranking posts by total engagement (or an “engagement rate”) to quickly identify the most effective content.
# Time-of-Day/Day-of-Week Insights
# Automated Summary or Insights
# Aggregate daily or hourly data to see when the page gets the most visitors or new followers.

# =============================================================================
# Load data from each file
# =============================================================================
competitor_data = load_excel_sheets(competitor_file)
visitors_data   = load_excel_sheets(visitors_file)
followers_data  = load_excel_sheets(followers_file)

# =============================================================================
# Sidebar: Data Source Selection
# =============================================================================
st.sidebar.image("gibs logo horiz_whitebackbluefont.png", width=200)
st.sidebar.header("BS4CL Page Metrics")
data_source = st.sidebar.radio("What would you like to analyse?", 
                               ["Competitor Analytics", "Followers", "Visitors"])

# Determine which dictionary of sheets to use based on selection.
if data_source == "Competitor Analytics":
    sheets = competitor_data
elif data_source == "Followers":
    sheets = followers_data
else:
    sheets = visitors_data

# Sidebar: Select a sheet from the chosen file.
if sheets:
    sheet_name = st.sidebar.selectbox("Select Sheet", list(sheets.keys()))
    df = sheets[sheet_name]
else:
    st.error("No sheets loaded from the file.")
    st.stop()

# =============================================================================
# Process Data Based on the Selected Data Source
# =============================================================================

# -------------------------------
# Competitor Analytics
# -------------------------------
if data_source == "Competitor Analytics":
    st.header("Competitor Analytics")
    try:
        # Assume the first row contains start_date and end_date info.
        start_date, end_date = df.iloc[0, 0], df.iloc[0, 1]
        df = df.iloc[1:]
        # The next row holds the column headers.
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)
    except Exception as e:
        st.error("Error during preprocessing: " + str(e))
    
    st.write(f"**Data Period:** {start_date} to {end_date}")
    st.subheader("Competitors")
    st.dataframe(df)
    
    # Automated Insights for key metrics.
    competitor_col = df.columns[0]
    automated_insights = []
    
    # Check for "Total Followers" metric.
    if "Total Followers" in df.columns:
        df["Total Followers"] = pd.to_numeric(df["Total Followers"], errors="coerce")
        max_idx = df["Total Followers"].idxmax()
        min_idx = df["Total Followers"].idxmin()
        automated_insights.append(
            f"The page with the highest followers is **{df.loc[max_idx, competitor_col]}** with **{df.loc[max_idx, 'Total Followers']}** followers, while **{df.loc[min_idx, competitor_col]}** has the lowest with **{df.loc[min_idx, 'Total Followers']}** followers."
        )
    # Check for "Total Posts" metric.
    if "Total posts" in df.columns:
        df["Total posts"] = pd.to_numeric(df["Total posts"], errors="coerce")
        max_idx = df["Total posts"].idxmax()
        min_idx = df["Total posts"].idxmin()
        automated_insights.append(
            f"The page with the highest posts is **{df.loc[max_idx, competitor_col]}** with **{df.loc[max_idx, 'Total posts']}** posts, while **{df.loc[min_idx, competitor_col]}** has the lowest with **{df.loc[min_idx, 'Total posts']}** posts."
        )
        
    if automated_insights:
        st.subheader("Insights")
        for insight in automated_insights:
            st.write(insight)
    else:
        st.info("No automated insights available as required metrics are missing.")
    
        
    # Interactive Bar Graph for Competitors
    st.subheader("Compare Metrics Across Competitors")
    # Assume the first column holds competitor names.
    competitor_col = df.columns[0]
    metric_options = df.columns[1:].tolist()

    # Allow users to select multiple metrics (default to first 4 if available)
    selected_metrics = st.multiselect("Select Metrics", metric_options, default=metric_options[:2])
        
    # Allow users to select competitors from the list.
    competitor_list = df[competitor_col].unique().tolist()
    default_competitors = ["Business Schools 4 Climate Leadership Africa", "Lagos Business School Sustainability Centre"]
    selected_competitors = st.multiselect("Select Competitors", competitor_list, default=default_competitors)
        
    filtered_df = df[df[competitor_col].isin(selected_competitors)]
    
    # Compute the sorting order for competitors based on the first selected metric (descending).
    sort_order = filtered_df.sort_values(by=selected_metrics[0], ascending=False)[competitor_col].tolist()

    # Reshape data from wide to long format for multiple metrics plotting.
    plot_df = filtered_df.melt(id_vars=competitor_col, value_vars=selected_metrics, 
                            var_name="Metric", value_name="Value")
        
    bar_chart = alt.Chart(plot_df).mark_bar().encode(
        x=alt.X(f"{competitor_col}:N", title="Competitor", sort=sort_order,
                axis=alt.Axis(
                    labelAngle=0,    # Keep labels horizontal
                    labelOverlap=True  # Let labels overlap or do 'greedy' to reduce collisions
        )),
        xOffset=alt.XOffset("Metric:N"),
        y=alt.Y("Value:Q", title="Value"),
        color=alt.Color("Metric:N", title="Metric"),
        tooltip=[competitor_col, "Metric:N", "Value:Q"]
    ).properties(
        width=700,
        height=400,
        title="Metrics Comparison Across Competitors"
    )
    st.altair_chart(bar_chart, use_container_width=True)


# -------------------------------
# Followers Analytics
# -------------------------------
if data_source == "Followers":
    st.header("Followers Analytics")
    # If the first row is all strings, assume it's a header.
    if all(isinstance(x, str) for x in df.iloc[0]):
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)
    
    # If there's no "Date" column, assume this is a non-time-series sheet.
    if 'Date' not in df.columns:
        # Convert columns with "total" in the heading to numeric.
        for col in df.columns:
            if "total" in str(col).lower():
                df[col] = pd.to_numeric(df[col], errors='coerce')
        st.dataframe(df)
        
        # Automated Insights for non–time-series data.
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        total_metrics = [col for col in numeric_cols if "total" in col.lower()]
        if total_metrics:
            for col in total_metrics:
                category_col = df.columns[0]  # Assume the first column holds the category.
                max_idx = df[col].idxmax()
                min_idx = df[col].idxmin()
                max_cat = df.loc[max_idx, category_col]
                min_cat = df.loc[min_idx, category_col]
                max_val = df.loc[max_idx, col]
                min_val = df.loc[min_idx, col]
                if max_val == min_val:
                    st.write(f"For **{col}**, all categories have the same value: **{max_val}**.")
                else:
                    st.write(f"For **{category_col}**, **{max_cat}** has the highest value (**{max_val}**) while **{min_cat}** has the lowest (**{min_val}**).")
        else:
            st.info("No numeric 'total' metrics available for automated insights.")
        
        
        # If numeric columns exist, allow a bar chart.
        if numeric_cols:
            selected_metric = st.selectbox("Select a Metric for Bar Chart", numeric_cols)
            # Use the first column as the category, if available.
            category_col = df.columns[0]
            # Sort order
            sort_order = df.sort_values(by=selected_metric, ascending=False)[category_col].tolist()

            bar_chart = alt.Chart(df).mark_bar().encode(
                x=alt.X(f"{category_col}:N", title=category_col, sort=sort_order),
                y=alt.Y(f"{selected_metric}:Q", title=selected_metric),
                tooltip=[category_col, selected_metric]
            ).properties(
                width=700,
                height=400,
                title=f"{selected_metric} by {category_col}"
            )
            st.altair_chart(bar_chart, use_container_width=True)
        st.stop()
    
    # Convert the Date column to datetime, the rest to numeric
    try:
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        
        for col in df.columns:
            if col != "Date":
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
    except Exception as e:
        st.error("Data type onversion error: " + str(e))
    
    # Date filters.
    min_date = df['Date'].min()
    max_date = df['Date'].max()
    start_date_filter = st.sidebar.date_input("Start Date", min_date, min_value=min_date, max_value=max_date)
    end_date_filter   = st.sidebar.date_input("End Date", max_date, min_value=min_date, max_value=max_date)
    
    if start_date_filter > end_date_filter:
        st.error("Error: Start Date must be before End Date.")
    else:
        mask = (df['Date'] >= pd.to_datetime(start_date_filter)) & (df['Date'] <= pd.to_datetime(end_date_filter))
        filtered_df = df.loc[mask].copy()
        st.subheader(f"New Followers between {start_date_filter} & {end_date_filter}")
        # st.dataframe(filtered_df)
        
        # Aggregate data by month.
        filtered_df['Date'] = filtered_df['Date'].dt.to_period('M').dt.to_timestamp()
        monthly_agg = filtered_df.groupby('Date').sum(numeric_only=True).reset_index()
        st.subheader("Monthly Aggregated Data")
        st.dataframe(monthly_agg)
        
        # Automated Summary for metrics with "total" in the header.
        # Find all numeric columns (other than 'Date') with "total" in their header.
        total_metrics = [
            col for col in monthly_agg.columns
            if col != "Date" and "total" in str(col).lower() and pd.api.types.is_numeric_dtype(monthly_agg[col])
        ]

        if total_metrics and len(monthly_agg) >= 2:
            for col in total_metrics:
                current_val = monthly_agg.iloc[-1][col]
                previous_val = monthly_agg.iloc[-2][col]
                if previous_val != 0:
                    pct_change = ((current_val - previous_val) / previous_val) * 100
                    current_month = monthly_agg.iloc[-1]['Date'].strftime('%b')
                    previous_month = monthly_agg.iloc[-2]['Date'].strftime('%b')
                    if pct_change > 0:
                        st.write(f"Great job! You gained {pct_change:.1f}% more {col} in {current_month} compared to {previous_month}.")
                    elif pct_change < 0:
                        st.write(f"A {abs(pct_change):.1f}% loss in {col} in {current_month} compared to {previous_month}.")
                    else:
                        st.write(f"There was no change in {col} from {previous_month} to {current_month}.")
                else:
                    st.write(f"No previous data to compare for {col}.")
        else:
            st.info("Not enough monthly data or no 'total' metrics available for summary.")

        
        # Exclude the Date column to get metric columns.
        metric_options = [col for col in monthly_agg.columns if col != "Date"]
        if not metric_options:
            st.error("No metric columns available for plotting.")
        else:
            # Allow the user to select one or more metrics.
            selected_metrics = st.multiselect("Select Metrics to Plot", metric_options, default=metric_options)
            if not selected_metrics:
                st.error("Please select at least one metric.")
            else:
                # Convert the DataFrame from wide to long format.
                plot_df = monthly_agg.melt(id_vars=['Date'], value_vars=selected_metrics, 
                                    var_name='Metric', value_name='Value')
                # Create an interval selection for zooming.
                zoom = alt.selection_interval(bind='scales', encodings=['x'])
                line_chart = alt.Chart(plot_df).mark_line(point=True).encode(
                    x=alt.X('Date:T', title="Date"),
                    y=alt.Y('Value:Q', title="Value"),
                    color=alt.Color('Metric:N', title="Metric"),
                    tooltip=['Date:T', 'Metric:N', 'Value:Q']
                ).properties(
                    width=700,
                    height=400,
                    title="Metrics Over Time"
                ).add_selection(zoom)
                st.altair_chart(line_chart, use_container_width=True)


# -------------------------------
# Visitors Analytics
# -------------------------------
if data_source == "Visitors":
    st.header("Visitors Analytics")
    if all(isinstance(x, str) for x in df.iloc[0]):
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)
    
    # If there's no "Date" column, assume this is a non-time-series sheet.
    if 'Date' not in df.columns:
        # Convert columns with "total" in the heading to numeric.
        for col in df.columns:
            if "total" in str(col).lower():
                df[col] = pd.to_numeric(df[col], errors='coerce')
        st.dataframe(df)
        
        # Automated Insights for non–time-series data.
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        total_metrics = [col for col in numeric_cols if "total" in col.lower()]
        if total_metrics:
            for col in total_metrics:
                category_col = df.columns[0]  # Assume the first column holds the category.
                max_idx = df[col].idxmax()
                min_idx = df[col].idxmin()
                max_cat = df.loc[max_idx, category_col]
                min_cat = df.loc[min_idx, category_col]
                max_val = df.loc[max_idx, col]
                min_val = df.loc[min_idx, col]
                if max_val == min_val:
                    st.write(f"For **{col}**, all categories have the same value: **{max_val}**.")
                else:
                    st.write(f"For **{category_col}**, **{max_cat}** has the highest value (**{max_val}**) while **{min_cat}** has the lowest (**{min_val}**).")
        else:
            st.info("No numeric 'total' metrics available for automated insights.")
        
        # If numeric columns exist, allow a bar chart.
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        
        if numeric_cols:
            selected_metric = st.selectbox("Select a Metric for Bar Chart", numeric_cols)
            # Use the first column as the category, if available.
            category_col = df.columns[0]
            sort_order = df.sort_values(by=selected_metric, ascending=False)[category_col].tolist()
            bar_chart = alt.Chart(df).mark_bar().encode(
                x=alt.X(f"{category_col}:N", title=category_col, sort=sort_order),
                y=alt.Y(f"{selected_metric}:Q", title=selected_metric),
                tooltip=[category_col, selected_metric]
            ).properties(
                width=700,
                height=400,
                title=f"{selected_metric} by {category_col}"
            )
            st.altair_chart(bar_chart, use_container_width=True)
        st.stop()
    
    # Convert the Date column to datetime, the rest to numeric
    try:
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        
        for col in df.columns:
            if col != "Date":
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
    except Exception as e:
        st.error("Data type onversion error: " + str(e))
    
    min_date = df['Date'].min()
    max_date = df['Date'].max()
    start_date_filter = st.sidebar.date_input("Start Date", min_date, min_value=min_date, max_value=max_date, key="visitors_start")
    end_date_filter   = st.sidebar.date_input("End Date", max_date, min_value=min_date, max_value=max_date, key="visitors_end")
    
    if start_date_filter > end_date_filter:
        st.error("Error: Start Date must be before End Date.")
    else:
        mask = (df['Date'] >= pd.to_datetime(start_date_filter)) & (df['Date'] <= pd.to_datetime(end_date_filter))
        filtered_df = df.loc[mask].copy()
        st.subheader("Filtered Data")
        # st.dataframe(filtered_df)
        
        # Aggregate data by month.
        filtered_df['Date'] = filtered_df['Date'].dt.to_period('M').dt.to_timestamp()
        monthly_agg = filtered_df.groupby('Date').sum(numeric_only=True).reset_index()
        st.subheader("Monthly Aggregated Data")
        st.dataframe(monthly_agg)
        
        # Automated Summary for metrics with "total" in the header.
        # Find all numeric columns (other than 'Date') with "total" in their header.
        total_metrics = [
            col for col in monthly_agg.columns
            if col != "Date" and "(total)" in str(col).lower() and pd.api.types.is_numeric_dtype(monthly_agg[col])
        ]

        if total_metrics and len(monthly_agg) >= 2:
            for col in total_metrics:
                current_val = monthly_agg.iloc[-1][col]
                previous_val = monthly_agg.iloc[-2][col]
                if previous_val != 0:
                    pct_change = ((current_val - previous_val) / previous_val) * 100
                    current_month = monthly_agg.iloc[-1]['Date'].strftime('%b')
                    previous_month = monthly_agg.iloc[-2]['Date'].strftime('%b')
                    if pct_change > 0:
                        st.write(f"Great job! You gained {pct_change:.1f}% more {col} in {current_month} compared to {previous_month}.")
                    elif pct_change < 0:
                        st.write(f"A {abs(pct_change):.1f}% loss in {col} in {current_month} compared to {previous_month}.")
                    else:
                        st.write(f"There was no change in {col} from {previous_month} to {current_month}.")
                else:
                    st.write(f"No previous data to compare for {col}.")
        else:
            st.info("Not enough monthly data or no 'total' metrics available for summary.")
        
        # Chart the data in a line chart with zooming functionality.
        agg_df = filtered_df.copy().sort_values("Date")
        # Exclude the Date column to get metric columns.
        metric_options = [col for col in monthly_agg.columns if col != "Date"]
        if not metric_options:
            st.error("No metric columns available for plotting.")
        else:
            # Allow the user to select one or more metrics.
            default_metrics = ["Total page views (total)", "Total unique visitors (total)", "Overview page views (total)"]
            selected_metrics = st.multiselect("Select Metrics to Plot", metric_options, default=default_metrics)
            if not selected_metrics:
                st.error("Please select at least one metric.")
            else:
                # Convert the DataFrame from wide to long format.
                plot_df = monthly_agg.melt(id_vars=['Date'], value_vars=selected_metrics, 
                                    var_name='Metric', value_name='Value')
                # Create an interval selection for zooming.
                zoom = alt.selection_interval(bind='scales', encodings=['x'])
                line_chart = alt.Chart(plot_df).mark_line(point=True).encode(
                    x=alt.X('Date:T', title="Date"),
                    y=alt.Y('Value:Q', title="Value"),
                    color=alt.Color('Metric:N', title="Metric"),
                    tooltip=['Date:T', 'Metric:N', 'Value:Q']
                ).properties(
                    width=700,
                    height=400,
                    title="Metrics Over Time"
                ).add_selection(zoom)
                st.altair_chart(line_chart, use_container_width=True)
