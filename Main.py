import os
from datetime import datetime
from nltk.sentiment import SentimentIntensityAnalyzer
import docx2txt
import docx
from docx.document import Document
from docx.opc.coreprops import CoreProperties
from wordcloud import WordCloud, STOPWORDS, ImageColorGenerator
import matplotlib.pyplot as plt
from datetime import datetime


class Wordy():
    def __init__(self):
        self.user_dir = ""
        self.abspath = ""
        self.cloud_dir = ""
        self.wordclouds = bool

    def user_input(self):
        """
        Function:
            Objective:
                Get user input from the user like directory name and whether to make wordclouds
            Procedure:
                Get directory from user and check if it exists
                If it does, ask if the user wants to generate wordclouds
                If yes, make a directory for them using datetime and cwd

        :return: None
        """
        while True:
            self.user_dir = input("Enter name of directory: ")
            if os.path.isdir(self.user_dir) is True:
                print("Directory check complete... directory valid")
                print("")

                cloud = input("Generate word cloud for files? (y/n) ")
                if cloud == "y":
                    self.wordclouds = True
                    cloud_dir = os.path.join(os.getcwd(), f"WordClouds - {datetime.today().strftime('%d-%m-%y)')}")
                    os.mkdir(cloud_dir)
                    self.cloud_dir = cloud_dir

                else:
                    self.wordclouds = False

                self.display()
                break

            else:
                print("Directory check complete... directory invalid\n")

    def display(self):
        """
        Function:
            Objective:
                Display the returns of other functions to the terminal
            Procedure:
                Count and print total number of files using OS
                Run each of the other fucntions and print their returns
                Check bool value before running word cloud generator function


        :return: None
        """
        count = 0
        for f in os.listdir(self.user_dir):
            count += 1
        print(f"Total files found in {self.user_dir}: {count}")

        print("File metadata:")
        print("------------------\n")

        for file in os.listdir(self.user_dir):
            self.abspath = os.path.abspath(os.path.join(self.user_dir, file))
            print(f"File: {file}")

            # Metadata using docx
            docx_meta = self.docx_metadata(self.abspath)
            for key, value in docx_meta.items():
                print(f"{key}: {value}")

            # Sentiment analysis here
            print("Sentiment analysis scores:")
            sa = self.get_sentiment(self.abspath)
            print(sa)

            print("")

            if self.wordclouds is True:
                self.cloud_gen(self.abspath)

    def get_sentiment(self, file):
        """
        Function:
            Objective:
                Use sentiment analysis on file
            Procedure:
            Instantiate sentiment analysis library
            Process text of file with library
            Get the score on processed text
            Return the score

        :param file: file object to work with
        :return: sentiment analysis score as dict
        """
        sia = SentimentIntensityAnalyzer()
        text = docx2txt.process(file)
        score = sia.polarity_scores(text)
        return score

    def docx_metadata(self, file):
        """
        Function:
            Objective:
                Get the metadata from a file
            Procedure:
                Instantiate docx library using file paths
                Set up dictionary
                Loop over different attributes and add to dictionary

        :param file: file to work on (not needed)
        :return: metadata - the dictionary of data
        """
        #path = self.abspath
        #os.chdir(path)

        #fname = file
        doc = docx.Document(self.abspath)

        prop = doc.core_properties

        metadata = {}
        for d in dir(prop):
            if not d.startswith('_'):
                metadata[d] = getattr(prop, d)

        return metadata

    def cloud_gen(self, file):
        """
        Function:
            Objective:
                Generate a wordcloud on a given file
            Procedure:
                Split file path to get name
                Process text from file
                Parse text to wordcloud library
                Save wordcloud as a png file named after file used

        :param file: file object to operate on
        :return: None
        """
        head, tail = os.path.split(file)
        text = docx2txt.process(file)
        wordcloud = WordCloud().generate(text)
        fname = os.path.join(self.cloud_dir, f"Wordcloud - {tail[:-5]}.png")
        wordcloud.to_file(fname)


c = Wordy()
c.user_input()
