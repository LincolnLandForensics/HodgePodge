import streamlit as st  # pip install streamlit
import pandas as pd
import numpy as np
 
# Title of the web app
st.title('Simple Streamlit App')

# Markdown can be used to create formatted text
st.markdown('''
This is a simple web app using Streamlit to demonstrate basic functionalities:
- Taking user inputs.
- Performing a calculation.
- Displaying results in various ways.
''')

# Sidebar for input
with st.sidebar:
    st.header('Input your values')
    # Asking user for input
    a = st.number_input('Enter first number', value=1)
    b = st.number_input('Enter second number', value=1)

# Perform a calculation
sum_ab = a + b

# Generate some data based on inputs
data = pd.DataFrame({
    'x': range(100),
    'y': np.random.normal(loc=sum_ab, scale=abs(a-b), size=100)
})

# Display the sum
st.subheader('Sum of the two numbers')
st.write('The sum of ', a, ' and ', b, ' is ', sum_ab, '.')

# Line chart to visualize the data
st.subheader('Line Chart of Generated Data')
st.line_chart(data.y)

# Display the DataFrame
st.subheader('Generated Data')
st.write(data)

# Optional: display raw code
st.subheader('Python Code')
st.code('''
import streamlit as st
import pandas as pd
import numpy as np

# Inputs
a = st.number_input('Enter first number', value=1)
b = st.number_input('Enter second number', value=1)

# Calculation
sum_ab = a + b

# Data generation
data = pd.DataFrame({
    'x': range(100),
    'y': np.random.normal(loc=sum_ab, scale=abs(a-b), size=100)
})

# Display
st.write('The sum of ', a, ' and ', b, ' is ', sum_ab, '.')
st.line_chart(data.y)
''')
