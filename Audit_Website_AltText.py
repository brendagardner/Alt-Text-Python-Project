#This Script will take URLs and pull the current images and alt text from the webpage and add them to an Excel file..
#created by Brenda Gardner, September 2025.
# Updated Dec 1 2025 to use alttext_utils helpers and reorganize columns and add text wrapping, skip logo images, scrape website for context.

# Step 1. Capture URLs and Subjects.
# a. Go to Website TOC and copy paste URLS.
# use in conjunction with https://slccbruins-my.sharepoint.com/:x:/g/personal/bgardn61_slcc_edu/IQD6AhzRzSxNRohDxUOKgeNYAQrWIctBAHw6ag8_bIYE-U8?e=sQFPWW

# End Step. After processing: Use macro to put the images in cell to allow filtering


import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse
from io import BytesIO
from pathlib import Path
from datetime import datetime
import warnings
from urllib3.exceptions import NotOpenSSLWarning
warnings.simplefilter("ignore", NotOpenSSLWarning)

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from helpers import alttext_utils as utils

# =========================
# CONFIG — EDIT THESE
# =========================
# These will be written into every row (can be left "" if you want to fill later).
DEFAULT_GRADE_BAND = "10-14"       # e.g. "10–12"
COURSE_TOPIC = "PreCalculus"      # e.g. "College Algebra"
BOOK_SOURCE_NAME = "OpenStax PreCalc 2E"

# Folder where the Excel output file will be written
OUTPUT_FOLDER = Path(
    "/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Sabbatical Files - Personal (Unshared)/1050 Information"
)

# List of (URL, Section/Topic) pairs
URLS_AND_TOPICS = [
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/01:_Functions/1.01:_Introduction_to_Functions", "1.1: Introduction to Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/01:_Functions/1.10:_Exercises", "1.10: Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/01:_Functions/1.10:_Exercises/1.10.01:_Review_Exercises", "1.10.1: Review Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/01:_Functions/1.10:_Exercises/1.10.02:_Practice_Test", "1.10.2: Practice Test"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/01:_Functions/1.02:_Functions_and_Function_Notation", "1.2: Functions and Function Notation"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/01:_Functions/1.03:_Domain_and_Range", "1.3: Domain and Range"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/01:_Functions/1.04:_Rates_of_Change_and_Behavior_of_Graphs", "1.4: Rates of Change and Behavior of Graphs"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/01:_Functions/1.05:_Composition_of_Functions", "1.5: Composition of Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/01:_Functions/1.06:_Transformation_of_Functions", "1.6: Transformation of Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/01:_Functions/1.07:_Absolute_Value_Functions", "1.7: Absolute Value Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/01:_Functions/1.08:_Inverse_Functions", "1.8: Inverse Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/01:_Functions/1.09:_Chapter_Review", "1.9: Chapter Review"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/01:_Functions/1.09:_Chapter_Review/1.9.01:_Key_Terms", "1.9.1: Key Terms"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/01:_Functions/1.09:_Chapter_Review/1.9.02:_Key_Equations", "1.9.2: Key Equations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/01:_Functions/1.09:_Chapter_Review/1.9.03:_Key_Concepts", "1.9.3: Key Concepts"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/10:_Analytic_Geometry", "10: Analytic Geometry"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/10:_Analytic_Geometry/10.01:_Introduction_to_Analytic_Geometry", "10.1: Introduction to Analytic Geometry"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/10:_Analytic_Geometry/10.02:_The_Ellipse", "10.2: The Ellipse"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/10:_Analytic_Geometry/10.03:_The_Hyperbola", "10.3: The Hyperbola"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/10:_Analytic_Geometry/10.04:_The_Parabola", "10.4: The Parabola"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/10:_Analytic_Geometry/10.05:_Rotation_of_Axes", "10.5: Rotation of Axes"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/10:_Analytic_Geometry/10.06:_Conic_Sections_in_Polar_Coordinates", "10.6: Conic Sections in Polar Coordinates"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/10:_Analytic_Geometry/10.07:_Chapter_Review", "10.7: Chapter Review"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/10:_Analytic_Geometry/10.07:_Chapter_Review/10.7.01:_Key_Terms", "10.7.1: Key Terms"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/10:_Analytic_Geometry/10.07:_Chapter_Review/10.7.02:_Key_Equations", "10.7.2: Key Equations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/10:_Analytic_Geometry/10.07:_Chapter_Review/10.7.03:_Key_Concepts", "10.7.3: Key Concepts"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/10:_Analytic_Geometry/10.08:_Exercises", "10.8: Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/10:_Analytic_Geometry/10.08:_Exercises/10.8.01:_Review_Exercises", "10.8.1: Review Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/10:_Analytic_Geometry/10.08:_Exercises/10.8.02:_Practice_Test", "10.8.2: Practice Test"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/10:_Analytic_Geometry/10.08:_Exercises/10.8.3:_Answers_to_Questions", "10.8.3: Answers to Questions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/11:_Sequences_Probability_and_Counting_Theory", "11: Sequences, Probability and Counting Theory"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/11:_Sequences_Probability_and_Counting_Theory/11.01:_Introduction_to_Sequences_Probability_and_Counting_Theory", "11.1: Introduction to Sequences, Probability and Counting Theory"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/11:_Sequences_Probability_and_Counting_Theory/11.10:_Exercises", "11.10: Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/11:_Sequences_Probability_and_Counting_Theory/11.10:_Exercises/11.10.01:_Review_Exercises", "11.10.1: Review Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/11:_Sequences_Probability_and_Counting_Theory/11.10:_Exercises/11.10.02:_Practice_Test", "11.10.2: Practice Test"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/11:_Sequences_Probability_and_Counting_Theory/11.02:_Sequences_and_Their_Notations", "11.2: Sequences and Their Notations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/11:_Sequences_Probability_and_Counting_Theory/11.03:_Arithmetic_Sequences", "11.3: Arithmetic Sequences"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/11:_Sequences_Probability_and_Counting_Theory/11.04:_Geometric_Sequences", "11.4: Geometric Sequences"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/11:_Sequences_Probability_and_Counting_Theory/11.05:_Series_and_Their_Notations", "11.5: Series and Their Notations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/11:_Sequences_Probability_and_Counting_Theory/11.06:_Counting_Principles", "11.6: Counting Principles"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/11:_Sequences_Probability_and_Counting_Theory/11.07:_Binomial_Theorem", "11.7: Binomial Theorem"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/11:_Sequences_Probability_and_Counting_Theory/11.08:_Probability", "11.8: Probability"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/11:_Sequences_Probability_and_Counting_Theory/11.09:_Chapter_Review", "11.9: Chapter Review"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/11:_Sequences_Probability_and_Counting_Theory/11.09:_Chapter_Review/11.9.01:_Key_Terms", "11.9.1: Key Terms"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/11:_Sequences_Probability_and_Counting_Theory/11.09:_Chapter_Review/11.9.02:_Key_Equations", "11.9.2: Key Equations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/11:_Sequences_Probability_and_Counting_Theory/11.09:_Chapter_Review/11.9.03:_Key_Concepts", "11.9.3: Key Concepts"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/12:_Introduction_to_Calculus", "12: Introduction to Calculus"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/12:_Introduction_to_Calculus/12.01:_Introduction_to_Calculus", "12.1: Introduction to Calculus"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/12:_Introduction_to_Calculus/12.02:_Finding_Limits_-_Numerical_and_Graphical_Approaches", "12.2: Finding Limits - Numerical and Graphical Approaches"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/12:_Introduction_to_Calculus/12.03:_Finding_Limits_-_Properties_of_Limits", "12.3: Finding Limits - Properties of Limits"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/12:_Introduction_to_Calculus/12.04:_Continuity", "12.4: Continuity"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/12:_Introduction_to_Calculus/12.05:_Derivatives", "12.5: Derivatives"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/12:_Introduction_to_Calculus/12.06:_Chapter_Review", "12.6: Chapter Review"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/12:_Introduction_to_Calculus/12.06:_Chapter_Review/12.6.01:_Key_Terms", "12.6.1: Key Terms"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/12:_Introduction_to_Calculus/12.06:_Chapter_Review/12.6.02:_Key_Equations", "12.6.2: Key Equations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/12:_Introduction_to_Calculus/12.06:_Chapter_Review/12.6.03:_Key_Concepts", "12.6.3: Key Concepts"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/12:_Introduction_to_Calculus/12.07:_Exercises", "12.7: Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/12:_Introduction_to_Calculus/12.07:_Exercises/12.7.01:_Review_Exercises", "12.7.1: Review Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/12:_Introduction_to_Calculus/12.07:_Exercises/12.7.02:_Practice_Test", "12.7.2: Practice Test"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/13:_Appendix", "13: Appendix"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/13:_Appendix/13.01:_Basic_Functions_and_Identities", "13.1: Basic Functions and Identities"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/13:_Appendix/13.02:_Answer_Key", "13.2: Answer Key"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/13:_Appendix/13.02:_Answer_Key/13.2.01:_Chapter_1", "13.2.1: Chapter 1"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/13:_Appendix/13.02:_Answer_Key/13.2.10:_Chapter_10", "13.2.10: Chapter 10"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/13:_Appendix/13.02:_Answer_Key/13.2.11:_Chapter_11", "13.2.11: Chapter 11"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/13:_Appendix/13.02:_Answer_Key/13.2.12:_Chapter_12", "13.2.12: Chapter 12"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/13:_Appendix/13.02:_Answer_Key/13.2.02:_Chapter_2", "13.2.2: Chapter 2"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/13:_Appendix/13.02:_Answer_Key/13.2.03:_Chapter_3", "13.2.3: Chapter 3"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/13:_Appendix/13.02:_Answer_Key/13.2.04:_Chapter_4", "13.2.4: Chapter 4"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/13:_Appendix/13.02:_Answer_Key/13.2.05:_Chapter_5", "13.2.5: Chapter 5"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/13:_Appendix/13.02:_Answer_Key/13.2.06:_Chapter_6", "13.2.6: Chapter 6"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/13:_Appendix/13.02:_Answer_Key/13.2.07:_Chapter_7", "13.2.7: Chapter 7"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/13:_Appendix/13.02:_Answer_Key/13.2.08:_Chapter_8", "13.2.8: Chapter 8"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/13:_Appendix/13.02:_Answer_Key/13.2.09:_Chapter_9", "13.2.9: Chapter 9"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/02:_Linear_Functions", "2: Linear Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/02:_Linear_Functions/2.01:_Introduction_to_Linear_Functions", "2.1: Introduction to Linear Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/02:_Linear_Functions/2.02:_Linear_Functions", "2.2: Linear Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/02:_Linear_Functions/2.03:_Graphs_of_Linear_Functions", "2.3: Graphs of Linear Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/02:_Linear_Functions/2.04:_Modeling_with_Linear_Functions", "2.4: Modeling with Linear Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/02:_Linear_Functions/2.05:_Fitting_Linear_Models_to_Data", "2.5: Fitting Linear Models to Data"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/02:_Linear_Functions/2.06:_Chapter_Review", "2.6: Chapter Review"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/02:_Linear_Functions/2.06:_Chapter_Review/2.6.01:_Key_Terms", "2.6.1: Key Terms"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/02:_Linear_Functions/2.06:_Chapter_Review/2.6.02:_Key_Equations", "2.6.2: Key Equations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/02:_Linear_Functions/2.06:_Chapter_Review/2.6.03:_Key_Concepts", "2.6.3: Key Concepts"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/02:_Linear_Functions/2.07:_Exercises", "2.7: Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/02:_Linear_Functions/2.07:_Exercises/2.7.01:_Review_Exercises", "2.7.1: Review Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/02:_Linear_Functions/2.07:_Exercises/2.7.02:_Practice_Test", "2.7.2: Practice Test"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/03:_Polynomial_and_Rational_Functions", "3: Polynomial and Rational Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/03:_Polynomial_and_Rational_Functions/3.01:_Introduction_to_Polynomial_and_Rational_Functions", "3.1: Introduction to Polynomial and Rational Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/03:_Polynomial_and_Rational_Functions/3.10:_Modeling_Using_Variation", "3.10: Modeling Using Variation"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/03:_Polynomial_and_Rational_Functions/3.11:_Chapter_Review", "3.11: Chapter Review"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/03:_Polynomial_and_Rational_Functions/3.11:_Chapter_Review/3.11.01:_Key_Terms", "3.11.1: Key Terms"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/03:_Polynomial_and_Rational_Functions/3.11:_Chapter_Review/3.11.02:_Key_Equations", "3.11.2: Key Equations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/03:_Polynomial_and_Rational_Functions/3.11:_Chapter_Review/3.11.03:_Key_Concepts", "3.11.3: Key Concepts"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/03:_Polynomial_and_Rational_Functions/3.12:_Exercises", "3.12: Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/03:_Polynomial_and_Rational_Functions/3.12:_Exercises/3.12.01:_Review_Exercises", "3.12.1: Review Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/03:_Polynomial_and_Rational_Functions/3.12:_Exercises/3.12.02:_Practice_Test", "3.12.2: Practice Test"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/03:_Polynomial_and_Rational_Functions/3.02:_Complex_Numbers", "3.2: Complex Numbers"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/03:_Polynomial_and_Rational_Functions/3.03:_Quadratic_Functions", "3.3: Quadratic Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/03:_Polynomial_and_Rational_Functions/3.04:_Power_Functions_and_Polynomial_Functions", "3.4: Power Functions and Polynomial Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/03:_Polynomial_and_Rational_Functions/3.05:_Graphs_of_Polynomial_Functions", "3.5: Graphs of Polynomial Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/03:_Polynomial_and_Rational_Functions/3.06:_Dividing_Polynomials", "3.6: Dividing Polynomials"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/03:_Polynomial_and_Rational_Functions/3.07:_Zeros_of_Polynomial_Functions", "3.7: Zeros of Polynomial Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/03:_Polynomial_and_Rational_Functions/3.08:_Rational_Functions", "3.8: Rational Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/03:_Polynomial_and_Rational_Functions/3.09:_Inverses_and_Radical_Functions", "3.9: Inverses and Radical Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/04:_Exponential_and_Logarithmic_Functions", "4: Exponential and Logarithmic Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/04:_Exponential_and_Logarithmic_Functions/4.01:_Introduction_to_Exponential_and_Logarithmic_Functions", "4.1: Introduction to Exponential and Logarithmic Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/04:_Exponential_and_Logarithmic_Functions/4.10:_Chapter_Review", "4.10: Chapter Review"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/04:_Exponential_and_Logarithmic_Functions/4.10:_Chapter_Review/4.10.01:_Key_Terms", "4.10.1: Key Terms"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/04:_Exponential_and_Logarithmic_Functions/4.10:_Chapter_Review/4.10.02:_Key_Equations", "4.10.2: Key Equations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/04:_Exponential_and_Logarithmic_Functions/4.10:_Chapter_Review/4.10.03:_Key_Concepts", "4.10.3: Key Concepts"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/04:_Exponential_and_Logarithmic_Functions/4.11:_Exercises", "4.11: Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/04:_Exponential_and_Logarithmic_Functions/4.11:_Exercises/4.11.01:_Review_Exercises", "4.11.1: Review Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/04:_Exponential_and_Logarithmic_Functions/4.11:_Exercises/4.11.02:_Practice_Test", "4.11.2: Practice Test"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/04:_Exponential_and_Logarithmic_Functions/4.02:_Exponential_Functions", "4.2: Exponential Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/04:_Exponential_and_Logarithmic_Functions/4.03:_Graphs_of_Exponential_Functions", "4.3: Graphs of Exponential Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/04:_Exponential_and_Logarithmic_Functions/4.04:_Logarithmic_Functions", "4.4: Logarithmic Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/04:_Exponential_and_Logarithmic_Functions/4.05:_Graphs_of_Logarithmic_Functions", "4.5: Graphs of Logarithmic Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/04:_Exponential_and_Logarithmic_Functions/4.06:_Logarithmic_Properties", "4.6: Logarithmic Properties"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/04:_Exponential_and_Logarithmic_Functions/4.07:_Exponential_and_Logarithmic_Equations", "4.7: Exponential and Logarithmic Equations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/04:_Exponential_and_Logarithmic_Functions/4.08:_Exponential_and_Logarithmic_Models", "4.8: Exponential and Logarithmic Models"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/04:_Exponential_and_Logarithmic_Functions/4.09:_Fitting_Exponential_Models_to_Data", "4.9: Fitting Exponential Models to Data"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/05:_Trigonometric_Functions", "5: Trigonometric Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/05:_Trigonometric_Functions/5.01:_Introduction_to_Trigonometric_Functions", "5.1: Introduction to Trigonometric Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/05:_Trigonometric_Functions/5.02:_Angles", "5.2: Angles"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/05:_Trigonometric_Functions/5.03:_Unit_Circle_-_Sine_and_Cosine_Functions", "5.3: Unit Circle - Sine and Cosine Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/05:_Trigonometric_Functions/5.04:_The_Other_Trigonometric_Functions", "5.4: The Other Trigonometric Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/05:_Trigonometric_Functions/5.05:_Right_Triangle_Trigonometry", "5.5: Right Triangle Trigonometry"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/05:_Trigonometric_Functions/5.06:_Chapter_Review", "5.6: Chapter Review"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/05:_Trigonometric_Functions/5.06:_Chapter_Review/5.6.01:_Key_Terms", "5.6.1: Key Terms"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/05:_Trigonometric_Functions/5.06:_Chapter_Review/5.6.02:_Key_Equations", "5.6.2: Key Equations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/05:_Trigonometric_Functions/5.06:_Chapter_Review/5.6.03:_Key_Concepts", "5.6.3: Key Concepts"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/05:_Trigonometric_Functions/5.07:_Exercises", "5.7: Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/05:_Trigonometric_Functions/5.07:_Exercises/5.7.01:_Review_Exercises", "5.7.1: Review Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/05:_Trigonometric_Functions/5.07:_Exercises/5.7.02:_Practice_Test", "5.7.2: Practice Test"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/06:_Periodic_Functions", "6: Periodic Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/06:_Periodic_Functions/6.01:_Introduction_to_Periodic_Functions", "6.1: Introduction to Periodic Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/06:_Periodic_Functions/6.02:_Graphs_of_the_Sine_and_Cosine_Functions", "6.2: Graphs of the Sine and Cosine Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/06:_Periodic_Functions/6.03:_Graphs_of_the_Other_Trigonometric_Functions", "6.3: Graphs of the Other Trigonometric Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/06:_Periodic_Functions/6.04:_Inverse_Trigonometric_Functions", "6.4: Inverse Trigonometric Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/06:_Periodic_Functions/6.05:_Chapter_Review", "6.5: Chapter Review"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/06:_Periodic_Functions/6.05:_Chapter_Review/6.5.01:_Key_Terms", "6.5.1: Key Terms"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/06:_Periodic_Functions/6.05:_Chapter_Review/6.5.02:_Key_Equations", "6.5.2: Key Equations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/06:_Periodic_Functions/6.05:_Chapter_Review/6.5.03:_Key_Concepts", "6.5.3: Key Concepts"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/06:_Periodic_Functions/6.06:_Exercises", "6.6: Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/06:_Periodic_Functions/6.06:_Exercises/6.6.01:_Review_Exercises", "6.6.1: Review Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/06:_Periodic_Functions/6.06:_Exercises/6.6.02:_Practice_Test", "6.6.2: Practice Test"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/07:_Trigonometric_Identities_and_Equations", "7: Trigonometric Identities and Equations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/07:_Trigonometric_Identities_and_Equations/7.01:_Introduction_to_Trigonometric_Identities_and_Equations", "7.1: Introduction to Trigonometric Identities and Equations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/07:_Trigonometric_Identities_and_Equations/7.02:_Solving_Trigonometric_Equations_with_Identities", "7.2: Solving Trigonometric Equations with Identities"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/07:_Trigonometric_Identities_and_Equations/7.03:_Sum_and_Difference_Identities", "7.3: Sum and Difference Identities"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/07:_Trigonometric_Identities_and_Equations/7.04:_Double-Angle_Half-Angle_and_Reduction_Formulas", "7.4: Double-Angle, Half-Angle, and Reduction Formulas"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/07:_Trigonometric_Identities_and_Equations/7.05:_Sum-to-Product_and_Product-to-Sum_Formulas", "7.5: Sum-to-Product and Product-to-Sum Formulas"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/07:_Trigonometric_Identities_and_Equations/7.06:_Solving_Trigonometric_Equations", "7.6: Solving Trigonometric Equations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/07:_Trigonometric_Identities_and_Equations/7.07:_Modeling_with_Trigonometric_Functions", "7.7: Modeling with Trigonometric Functions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/07:_Trigonometric_Identities_and_Equations/7.08:_Chapter_Review", "7.8: Chapter Review"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/07:_Trigonometric_Identities_and_Equations/7.08:_Chapter_Review/7.8.01:_Key_Terms", "7.8.1: Key Terms"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/07:_Trigonometric_Identities_and_Equations/7.08:_Chapter_Review/7.8.02:_Key_Equations", "7.8.2: Key Equations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/07:_Trigonometric_Identities_and_Equations/7.08:_Chapter_Review/7.8.03:_Key_Concepts", "7.8.3: Key Concepts"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/07:_Trigonometric_Identities_and_Equations/7.09:_Exercises", "7.9: Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/07:_Trigonometric_Identities_and_Equations/7.09:_Exercises/7.9.01:_Review_Exercises", "7.9.1: Review Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/07:_Trigonometric_Identities_and_Equations/7.09:_Exercises/7.9.02:_Practice_Test", "7.9.2: Practice Test"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/08:_Further_Applications_of_Trigonometry", "8: Further Applications of Trigonometry"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/08:_Further_Applications_of_Trigonometry/8.01:_Introduction_to_Further_Applications_of_Trigonometry", "8.1: Introduction to Further Applications of Trigonometry"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/08:_Further_Applications_of_Trigonometry/8.10:_Chapter_Review", "8.10: Chapter Review"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/08:_Further_Applications_of_Trigonometry/8.10:_Chapter_Review/8.10.01:_Key_Terms", "8.10.1: Key Terms"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/08:_Further_Applications_of_Trigonometry/8.10:_Chapter_Review/8.10.02:_Key_Equations", "8.10.2: Key Equations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/08:_Further_Applications_of_Trigonometry/8.10:_Chapter_Review/8.10.03:_Key_Concepts", "8.10.3: Key Concepts"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/08:_Further_Applications_of_Trigonometry/8.11:_Exercises", "8.11: Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/08:_Further_Applications_of_Trigonometry/8.11:_Exercises/8.11.01:_Review_Exercises", "8.11.1: Review Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/08:_Further_Applications_of_Trigonometry/8.11:_Exercises/8.11.02:_Practice_Test", "8.11.2: Practice Test"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/08:_Further_Applications_of_Trigonometry/8.02:_Non-right_Triangles_-_Law_of_Sines", "8.2: Non-right Triangles - Law of Sines"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/08:_Further_Applications_of_Trigonometry/8.03:_Non-right_Triangles_-_Law_of_Cosines", "8.3: Non-right Triangles - Law of Cosines"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/08:_Further_Applications_of_Trigonometry/8.04:_Polar_Coordinates", "8.4: Polar Coordinates"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/08:_Further_Applications_of_Trigonometry/8.05:_Polar_Coordinates_-_Graphs", "8.5: Polar Coordinates - Graphs"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/08:_Further_Applications_of_Trigonometry/8.06:_Polar_Form_of_Complex_Numbers", "8.6: Polar Form of Complex Numbers"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/08:_Further_Applications_of_Trigonometry/8.07:_Parametric_Equations", "8.7: Parametric Equations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/08:_Further_Applications_of_Trigonometry/8.08:_Parametric_Equations_-_Graphs", "8.8: Parametric Equations - Graphs"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/08:_Further_Applications_of_Trigonometry/8.09:_Vectors", "8.9: Vectors"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/09:_Systems_of_Equations_and_Inequalities", "9: Systems of Equations and Inequalities"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/09:_Systems_of_Equations_and_Inequalities/9.01:_Introduction_to_Systems_of_Equations_and_Inequalities", "9.1: Introduction to Systems of Equations and Inequalities"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/09:_Systems_of_Equations_and_Inequalities/9.10:_Chapter_Review", "9.10: Chapter Review"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/09:_Systems_of_Equations_and_Inequalities/9.10:_Chapter_Review/9.10.01:_Key_Terms", "9.10.1: Key Terms"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/09:_Systems_of_Equations_and_Inequalities/9.10:_Chapter_Review/9.10.02:_Key_Equations", "9.10.2: Key Equations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/09:_Systems_of_Equations_and_Inequalities/9.10:_Chapter_Review/9.10.03:_Key_Concepts", "9.10.3: Key Concepts"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/09:_Systems_of_Equations_and_Inequalities/9.11:_Exercises", "9.11: Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/09:_Systems_of_Equations_and_Inequalities/9.11:_Exercises/9.11.01:_Review_Exercises", "9.11.1: Review Exercises"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/09:_Systems_of_Equations_and_Inequalities/9.11:_Exercises/9.11.02:_Practice_Test", "9.11.2: Practice Test"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/09:_Systems_of_Equations_and_Inequalities/9.02:_Systems_of_Linear_Equations_-_Two_Variables", "9.2: Systems of Linear Equations - Two Variables"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/09:_Systems_of_Equations_and_Inequalities/9.03:_Systems_of_Linear_Equations_-_Three_Variables", "9.3: Systems of Linear Equations - Three Variables"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/09:_Systems_of_Equations_and_Inequalities/9.04:_Systems_of_Nonlinear_Equations_and_Inequalities_-_Two_Variables", "9.4: Systems of Nonlinear Equations and Inequalities - Two Variables"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/09:_Systems_of_Equations_and_Inequalities/9.05:_Partial_Fractions", "9.5: Partial Fractions"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/09:_Systems_of_Equations_and_Inequalities/9.06:_Matrices_and_Matrix_Operations", "9.6: Matrices and Matrix Operations"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/09:_Systems_of_Equations_and_Inequalities/9.07:_Solving_Systems_with_Gaussian_Elimination", "9.7: Solving Systems with Gaussian Elimination"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/09:_Systems_of_Equations_and_Inequalities/9.08:_Solving_Systems_with_Inverses", "9.8: Solving Systems with Inverses"),
("https://math.libretexts.org/Bookshelves/Precalculus/Precalculus_2e_(OpenStax)/09:_Systems_of_Equations_and_Inequalities/9.09:_Solving_Systems_with_Cramer's_Rule", "9.9: Solving Systems with Cramer's Rule"),
]


class CombinedImageExtractor:
    def __init__(self, proxies=None):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': (
                'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                'AppleWebKit/537.36 (KHTML, like Gecko) '
                'Chrome/91.0.4472.124 Safari/537.36'
            )
        })
        if proxies:
            self.session.proxies.update(proxies)

    def extract_images_from_url(self, url, topic=""):
        """Extract all images and alt text from a webpage and its iframes."""
        try:
            if not url:
                return []

            print(f"Processing main page: {url}")
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')

            images_data = []

            # Extract images from main page
            main_images = self._extract_images_from_soup(
                soup, url, topic, source_type="main page"
            )
            images_data.extend(main_images)

            # Find and process all iframes
            iframes = soup.find_all('iframe')
            print(f"Found {len(iframes)} iframes to process")

            for i, iframe in enumerate(iframes):
                iframe_src = iframe.get('src')
                if not iframe_src:
                    continue

                # Convert relative URLs to absolute
                iframe_url = urljoin(url, iframe_src)

                # Skip certain iframe types that won't have useful images
                skip_domains = [
                    'youtube.com', 'vimeo.com', 'facebook.com',
                    'twitter.com', 'ads.', 'analytics.', 'google.com/maps'
                ]
                if any(domain in iframe_url.lower() for domain in skip_domains):
                    print(f"Skipping iframe {i + 1}: {iframe_url} "
                          f"(likely ads/social/video)")
                    continue

                print(f"Processing iframe {i + 1}: {iframe_url}")
                iframe_images = self._extract_images_from_iframe(
                    iframe_url, topic, original_webpage=url
                )
                images_data.extend(iframe_images)

            print(f"Total found: {len(images_data)} images")
            return images_data

        except Exception as e:
            print(f"Error extracting from {url}: {str(e)}")
            return []

    def _extract_images_from_soup(
            self,
            soup,
            source_url,
            topic,
            source_type="page",
            original_webpage=None
    ):
        """Helper method to extract images from a BeautifulSoup object."""
        images_data = []
        img_tags = soup.find_all('img')

        for img in img_tags:
            src = img.get('src')
            if not src:
                continue

            # --------------------------------------------------------
            # 1️⃣ Skip images inside <header> or <footer>
            # --------------------------------------------------------
            if img.find_parent(['header', 'footer']):
                continue

            # --------------------------------------------------------
            # 2️⃣ Skip by parent class/id (common footer/header patterns)
            # --------------------------------------------------------
            parent = img.parent
            if parent:
                skip_classes = [
                    "site-header", "header", "footer", "site-footer",
                    "footer-logo", "masthead", "page-header"
                ]
                skip_ids = ["header", "footer", "colophon", "masthead"]

                # skip if parent contains a class we are avoiding
                if any(cls in skip_classes for cls in parent.get("class", [])):
                    continue

                # skip if parent has an ID we are avoiding
                if parent.get("id") in skip_ids:
                    continue

            # --------------------------------------------------------
            # 3️⃣ Skip by URL filename patterns (logos, banners, icons)
            # --------------------------------------------------------
            # Convert relative URLs to absolute NOW so we can inspect filename
            full_url = urljoin(source_url, src)
            full_url_lower = full_url.lower()

            skip_patterns = [
                "logo", "banner", "footer", "header",
                "favicon", "sprite", "icon"
            ]

            if any(p in full_url_lower for p in skip_patterns):
                continue

            # --------------------------------------------------------
            # SAFE — process the content image
            # --------------------------------------------------------

            alt_text = img.get('alt', '') or ""
            title_text = img.get('title', '') or ""
            author = img.get('data-author') or img.get('data-credit') or ''
            copyright_text = img.get('data-copyright') or img.get('copyright') or ''

            # grab context around the image
            image_context = self._extract_image_context(img)

            images_data.append({
                'image_location': full_url,
                'alt_text': alt_text,
                'webpage': original_webpage if original_webpage else source_url,
                'section': topic,
                'image_title': title_text,
                'author': author,
                'copyright': copyright_text,
                'source_type': source_type,
                'image_context': image_context,  # 🔹 NEW FIELD
            })

        print(f"Found {len(images_data)} images in {source_type}")
        return images_data

    def _extract_image_context(self, img, max_chars: int = 500) -> str:
        """
        Try to pull a meaningful text context for an <img> element.

        Priority:
          1) <figcaption> (if inside a <figure>)
          2) Closest previous sibling <p>/<li>/<div> with real text
          3) Walk up ancestors, checking their previous siblings
          4) Fallback: closest following <p>/<li>/<div> with text
        """
        def clean_text(node):
            if not node:
                return ""
            txt = node.get_text(" ", strip=True)
            txt = " ".join(txt.split())
            return txt[:max_chars]

        # 1) If inside a <figure>, prefer <figcaption>
        figure = img.find_parent("figure")
        if figure:
            cap = figure.find("figcaption")
            if cap:
                cap_text = clean_text(cap)
                if cap_text:
                    return cap_text

        # Helper to scan previous siblings of a node
        def prev_sibling_text(node):
            for sib in getattr(node, "previous_siblings", []):
                if getattr(sib, "name", None) in ("p", "li", "div"):
                    txt = clean_text(sib)
                    if txt:
                        return txt
            return ""

        # 2) Previous siblings of the <img> itself
        ctx = prev_sibling_text(img)
        if ctx:
            return ctx

        # 3) Walk up ancestors and check THEIR previous siblings
        parent = img.parent
        while parent is not None and getattr(parent, "name", None) not in ("body", "html"):
            ctx = prev_sibling_text(parent)
            if ctx:
                return ctx
            parent = parent.parent

        # 4) Fallback: next siblings after the <img>
        for sib in getattr(img, "next_siblings", []):
            if getattr(sib, "name", None) in ("p", "li", "div"):
                txt = clean_text(sib)
                if txt:
                    return txt

        return ""

    def _extract_images_from_iframe(self, iframe_url, topic, original_webpage):
        """Extract images from an iframe."""
        try:
            response = self.session.get(iframe_url, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')

            return self._extract_images_from_soup(
                soup, iframe_url, topic, source_type=f"iframe ({iframe_url})",
                original_webpage=original_webpage
            )

        except Exception as e:
            print(f"Error processing iframe {iframe_url}: {str(e)}")
            return []

    def _safe_filename(self, name: str) -> str:
        return "".join(
            c if c.isalnum() or c in (' ', '-', '_') else '_'
            for c in name
        ).strip().replace(' ', '_')

    def _extract_book_name_from_url(self, url: str) -> str:
        """Extract first path segment as the book name, e.g., slcc-elementaryalgebra."""
        try:
            parsed = urlparse(url)
            first_segment = (parsed.path or '/').strip('/').split('/')[0] or 'book'
            return self._safe_filename(first_segment)
        except Exception:
            return "book"

    # ------------------------------------------------------------------
    # NEW: Excel builder using alttext_utils + new column layout
    # ------------------------------------------------------------------
    def create_excel_with_images(
        self,
        all_images_data,
        output_filename: str,
        file_source: str = "",
        grade_band: str = "",
        course_topic: str = ""
    ):
        """
        Create a single Excel file with embedded thumbnails and the
        MASTER 30-column layout (A–AD), standardized across all scripts:

        A:  Image Thumbnail
        B:  Image Filename
        C:  Picture ID          — blank for website sources
        D:  Image URL           (hyperlink to image)
        E:  Image Size
        F:  Image Topic
        G:  Image Context
        H:  Page #              — blank for website sources
        I:  Section/Module/Figure #
        J:  Filename (Hyperlinked) or Webpage Name (URL)
        K:  Grade Level
        L:  Course Name
        M:  Original Source Author
        N:  Copyright
        O:  Date of Audit       — auto-filled
        P:  Original Alt Text
        Q:  Original Alt Text Character Count  — =LEN(P)
        R:  Original Alt Text Notes            — blank
        S:  A.I. Short Alt Text                — blank
        T:  A.I. Med Alt Text                  — blank
        U:  A.I. Long Alt Text                 — blank
        V:  A.I. Notes                         — blank
        W:  New/Suggested Alt Text             — blank
        X:  New/Suggested Alt Text Count       — =LEN(W)
        Y:  New/Suggested Alt Text Notes       — blank
        Z:  Status - Alt Text Completed?       — blank
        AA: Reviewer Notes                     — blank
        AB: In Word?                           — blank for website sources
        AC: In Excel?                          — blank
        AD: Date of Completion                 — blank
        """
        if not all_images_data:
            print("No images to export.")
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Extracted Images"

        # ---- Headers — Master 30-column layout A–AD ----
        headers = [
            "Image Thumbnail",                               # A  (1)
            "Image Filename",                                # B  (2)
            "Picture ID",                                    # C  (3) — blank for website sources
            "Image URL",                                     # D  (4)
            "Image Size",                                    # E  (5)
            "Image Topic",                                   # F  (6)
            "Image Context",                                 # G  (7)
            "Page #",                                        # H  (8) — blank for website sources
            "Section/Module/Figure #",                       # I  (9)
            "Filename (Hyperlinked) or Webpage Name (URL)",  # J  (10)
            "Grade Level",                                   # K  (11)
            "Course Name",                                   # L  (12)
            "Original Source Author",                        # M  (13)
            "Copyright",                                     # N  (14)
            "Date of Audit",                                 # O  (15) — auto-filled
            "Original Alt Text",                             # P  (16)
            "Original Alt Text Character Count",             # Q  (17) — =LEN(P)
            "Original Alt Text Notes",                       # R  (18) — blank
            "A.I. Short Alt Text",                           # S  (19) — blank
            "A.I. Med Alt Text",                             # T  (20) — blank
            "A.I. Long Alt Text",                            # U  (21) — blank
            "A.I. Notes",                                    # V  (22) — blank
            "New/Suggested Alt Text",                        # W  (23) — blank
            "New/Suggested Alt Text Count",                  # X  (24) — =LEN(W)
            "New/Suggested Alt Text Notes",                  # Y  (25) — blank
            "Status - Alt Text Completed?\n(Original/Suggested/N)",  # Z  (26) — blank
            "Reviewer Notes",                                # AA (27) — blank
            "In Word?",                                      # AB (28) — blank for website sources
            "In Excel?",                                     # AC (29) — blank
            "Date of Completion",                            # AD (30) — blank
        ]

        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = Font(bold=True)

        # Column widths — Master 30-column layout
        ws.column_dimensions['A'].width = utils._pixels_to_column_width(600)  # Image Thumbnail
        ws.column_dimensions['B'].width = 30   # Image Filename
        ws.column_dimensions['C'].width = 18   # Picture ID
        ws.column_dimensions['D'].width = 40   # Image URL
        ws.column_dimensions['E'].width = 15   # Image Size
        ws.column_dimensions['F'].width = 25   # Image Topic
        ws.column_dimensions['G'].width = 40   # Image Context
        ws.column_dimensions['H'].width = 8    # Page #
        ws.column_dimensions['I'].width = 25   # Section/Module/Figure #
        ws.column_dimensions['J'].width = 40   # Filename (Hyperlinked) or Webpage Name
        ws.column_dimensions['K'].width = 12   # Grade Level
        ws.column_dimensions['L'].width = 20   # Course Name
        ws.column_dimensions['M'].width = 20   # Original Source Author
        ws.column_dimensions['N'].width = 20   # Copyright
        ws.column_dimensions['O'].width = 16   # Date of Audit
        ws.column_dimensions['P'].width = 50   # Original Alt Text
        ws.column_dimensions['Q'].width = 16   # Original Alt Text Character Count
        ws.column_dimensions['R'].width = 40   # Original Alt Text Notes
        ws.column_dimensions['S'].width = 50   # A.I. Short Alt Text
        ws.column_dimensions['T'].width = 50   # A.I. Med Alt Text
        ws.column_dimensions['U'].width = 50   # A.I. Long Alt Text
        ws.column_dimensions['V'].width = 40   # A.I. Notes
        ws.column_dimensions['W'].width = 50   # New/Suggested Alt Text
        ws.column_dimensions['X'].width = 18   # New/Suggested Alt Text Count
        ws.column_dimensions['Y'].width = 40   # New/Suggested Alt Text Notes
        ws.column_dimensions['Z'].width = 20   # Status - Alt Text Completed?
        ws.column_dimensions['AA'].width = 30  # Reviewer Notes
        ws.column_dimensions['AB'].width = 12  # In Word?
        ws.column_dimensions['AC'].width = 12  # In Excel?
        ws.column_dimensions['AD'].width = 18  # Date of Completion

        # ---- Data rows ----
        current_row = 2

        for img_info in all_images_data:
            image_url = img_info.get("image_location", "") or ""
            alt_text = img_info.get("alt_text", "") or ""
            webpage = img_info.get("webpage", "") or ""
            section = img_info.get("section", "") or ""
            image_title = img_info.get("image_title", "") or ""
            author = img_info.get("author", "") or ""
            copyright_text = img_info.get("copyright", "") or ""
            image_context = img_info.get("image_context", "") or ""  # 🔹 NEW

            # B: image filename from URL path
            parsed = urlparse(image_url)
            image_filename = Path(parsed.path).name if parsed.path else ""

            # K/L: grade level & course topic pulled from config
            grade_value = grade_band or ""
            course_value = course_topic or ""

            # P: Original Alt Text
            current_alt = alt_text

            # Column B (2): Image Filename
            ws.cell(row=current_row, column=2, value=image_filename)

            # Column C (3): Picture ID — blank for website sources

            # Column D (4): Image URL (hyperlink)
            ws.cell(row=current_row, column=4, value=image_url)
            if image_url:
                ws.cell(row=current_row, column=4).hyperlink = image_url
                ws.cell(row=current_row, column=4).style = "Hyperlink"

            # Column F (6): Image Topic (using HTML title as a placeholder)
            ws.cell(row=current_row, column=6, value=image_title)

            # Column G (7): Image Context (paragraph before the image)
            ws.cell(row=current_row, column=7, value=image_context)

            # Column H (8): Page # — blank for website sources

            # Column I (9): Section/Module/Figure #
            ws.cell(row=current_row, column=9, value=section)

            # Column J (10): Webpage/Source (hyperlink)
            ws.cell(row=current_row, column=10, value=webpage)
            if webpage:
                ws.cell(row=current_row, column=10).hyperlink = webpage
                ws.cell(row=current_row, column=10).style = "Hyperlink"

            # Column K (11): Grade Level
            ws.cell(row=current_row, column=11, value=grade_value)

            # Column L (12): Course Name
            ws.cell(row=current_row, column=12, value=course_value)

            # Column M (13): Original Source Author
            ws.cell(row=current_row, column=13, value=author)

            # Column N (14): Copyright
            ws.cell(row=current_row, column=14, value=copyright_text)

            # Column O (15): Date of Audit — auto-filled with today's date
            ws.cell(row=current_row, column=15, value=datetime.now().strftime("%m/%d/%Y"))

            # Column P (16): Original Alt Text
            ws.cell(row=current_row, column=16, value=current_alt)

            # Column Q (17): Original Alt Text Character Count — formula =LEN(P)
            ws.cell(row=current_row, column=17, value=f"=LEN(P{current_row})")

            # Columns R–V (18–22): Original Alt Text Notes, A.I. columns — blank
            # Column W (23): New/Suggested Alt Text — blank

            # Column X (24): New/Suggested Alt Text Count — formula =LEN(W)
            ws.cell(row=current_row, column=24, value=f"=LEN(W{current_row})")

            # Columns Y–AA (25–27): Notes / Status / Reviewer Notes — blank
            # Column AB (28): In Word? — blank for website sources
            # Columns AC–AD (29–30): In Excel? / Date of Completion — blank

            # --- A: Thumbnail + E: Image Size ---
            try:
                if image_url:
                    print(f"Downloading image for row {current_row}: {image_url}")
                    resp = self.session.get(image_url, timeout=10)
                    resp.raise_for_status()

                    # Use shared thumbnail helper
                    thumb_bytes, thumb_w, thumb_h, orig_w, orig_h = utils.create_thumbnail_from_bytes(
                        resp.content,
                        max_width_px=600,
                        allow_upscale=False,
                    )

                    # E: Image Size (original dimensions)
                    if orig_w and orig_h:
                        img_size_str = f"{orig_w} x {orig_h} px"
                        ws.cell(row=current_row, column=5, value=img_size_str)

                    # Insert thumbnail into Column A
                    if thumb_bytes:
                        utils.insert_png_bytes_in_cell(
                            ws,
                            row=current_row,
                            col=1,  # Column A
                            png_bytes=thumb_bytes,
                            width_px=thumb_w,
                            height_px=thumb_h,
                            adjust_row_col=True,
                            max_col_width_px=600,
                        )
                    else:
                        ws.row_dimensions[current_row].height = 20
                else:
                    # no URL, no thumbnail
                    ws.row_dimensions[current_row].height = 20

            except Exception as e:
                print(f"Failed to insert thumbnail at row {current_row}: {e}")
                ws.row_dimensions[current_row].height = 20

            current_row += 1


        # Wrap text and autofit B–AD (all text columns), cap at ~240px
        text_columns = [
            "B", "C", "D", "E", "F", "G",
            "H", "I", "J", "K", "L", "M",
            "N", "O", "P", "Q", "R", "S",
            "T", "U", "V", "W", "X", "Y",
            "Z", "AA", "AB", "AC", "AD",
        ]
        utils.wrap_and_autofit_columns(
            ws,
            text_columns=text_columns,
            max_col_width_px=240,
        )

        # AUTO-ADJUST ROW HEIGHTS
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            max_lines = 1
            for cell in row:
                if cell.value is not None:
                    text = str(cell.value)
                    # Rough estimate: every ~60 characters creates a new wrapped line
                    est_lines = max(1, len(text) // 60 + 1)
                    if est_lines > max_lines:
                        max_lines = est_lines

            # Convert estimated lines into row height
            ws.row_dimensions[cell.row].height = max_lines * 30

        wb.save(output_filename)
        print(f"Excel file with embedded images saved as: {output_filename}")


def main():
    extractor = CombinedImageExtractor()

    all_images_data = []

    for url, topic in URLS_AND_TOPICS:
        if not url:
            continue
        images_data = extractor.extract_images_from_url(url, topic)
        all_images_data.extend(images_data)

    if not all_images_data:
        print("No images found to process.")
        return

    date_str = datetime.now().strftime("%Y%m%d")
    output_filename = OUTPUT_FOLDER / f"{BOOK_SOURCE_NAME}_website_{date_str}.xlsx"

    print(f"\nCreating Excel file with {len(all_images_data)} total images...")
    extractor.create_excel_with_images(
        all_images_data,
        output_filename,
        file_source=BOOK_SOURCE_NAME,
        grade_band=DEFAULT_GRADE_BAND,
        course_topic=COURSE_TOPIC
    )

    print(f"Done! Excel saved to:\n{output_filename}")


if __name__ == "__main__":
    main()
