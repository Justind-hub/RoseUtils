import pandas as pd
import traceback


def run(self, output, comments):
    try:
        STORELIST = [1740, 1743, 2172, 2174, 2236, 2272, 2457, 2603, 2549, 2953, 3498, 4778]
        KEEP = ["Why Highly Satsfied","FoodTrac tell us more","Compliment"]
        df = pd.read_csv(comments, skiprows=5, usecols=[1,2,3,5,6,7,8,9,10,18,19,21,22])
        df = df.reset_index()

        def test(row):
            if row['Unit'] not in STORELIST or row['Open End'] not in KEEP or row['Overall Satisfaction'] < 3:
                df.drop(row['index'], inplace = True)

        df.apply(test, axis=1)
        df.to_html(output+"Comments.html", na_rep="",columns=['Unit','Feedback Date','Event Date','Comment','Manager In Charge','Delivery Driver'],index=False)
        self.ui.outputbox.setText(f"Report outputed to \n{output}Comments.html")
    except:
        self.ui.outputbox.append("ENCOUNTERED ERROR")
        self.ui.outputbox.append("Please send the contents of this box to Justin")
        self.ui.outputbox.append(traceback.format_exc())








if __name__ == '__main__':
    run("C:/Users/justi/OneDrive/Desktop/", "C:/Users/justi/Downloads/2022-08-15_2022-08-28_Comments_ByComment.csv")

