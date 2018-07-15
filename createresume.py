import uno
import datetime
import time

JOB_DETAILS1 = {
    'employer': 'The Startup Company',
    'title': 'Poobah',
    'city': 'Seattle',
    'location2': 'WA',
    'start': datetime.date(2014, 7, 1),
    'end': datetime.date(2018, 4, 1),
    'duties': 'Founded a company and executed a high-growth strategy',
    'accomplishments': [
        'grew by 98% first year',
        'retained 70% of employees over three year period',
        'negotiated an acquisition netting investors a 35-fold return',
    ]
}

class ResumeCreator:

    def set_up(self):
        global smgr
        localContext = uno.getComponentContext()
        self.resolver = localContext.ServiceManager.createInstanceWithContext(
                        "com.sun.star.bridge.UnoUrlResolver", localContext )
        self.context = self.connect_to_running_office()
        self.service_manager = self.context.ServiceManager
        self.desktop = self.service_manager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.context)
        self.model = self.get_current_writer_doc()

    def connect_to_running_office(self):
        return self.resolver.resolve(
            "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")

    def do_set_up(self):
        self.set_up()
        return {
            'resolver': self.resolver,
            'context': self.context,
            'desktop': self.desktop,
            'model': self.model,
        }

    def get_current_writer_doc(self):
        self.current_writer_doc = self.desktop.getCurrentComponent()
        return self.current_writer_doc

    def create_table(self, rows, columns):
        TABLE_STRING = "com.sun.star.text.TextTable"  # https://wiki.openoffice.org/wiki/Writer/API/Tables
        the_table = self.current_writer_doc.createInstance(TABLE_STRING)
        the_table.initialize(rows, columns)
        return the_table

    def create_new_file(self):
        # https://wiki.openoffice.org/wiki/Writer/API/Overview#Creating.2C_opening_a_Writer_Document
        SWRITER = "private:factory/swriter"
        BLANK = "_blank"
        FALSE = 0
        try:
            newdoc = self.desktop.loadComponentFromURL(
                SWRITER, BLANK, FALSE, [])
        except:
            raise RuntimeWarning("Creating a new doc didn't work.")
        self.current_writer_doc = newdoc
        return newdoc

    def insert_job(self, job_details, current_doc=None):
        self.current_writer_doc = current_doc or self.current_writer_doc
        table = self.insert_table_for_job(job_details)
        paragraph = self.insert_paragraph_for_job(job_details)

    def insert_paragraph_for_job(self, job_details):
        text = self.current_writer_doc.Text
        cursor = text.createTextCursor()
        for bullet_point in job_details['accomplishments']:
            text.insertString(cursor, "* ", 0)
            text.insertString(cursor, bullet_point, 0)
            text.insertString(cursor, "\n", 0)

    def insert_table(self, table):
        text = self.current_writer_doc.Text
        cursor = text.createTextCursor()
        text.insertTextContent(cursor, table, 0)

    def insert_table_for_job(self, job_details):
        table = self.create_table(1, 2)
        self.insert_table(table)
        jd = job_details
        left_side = "{}, {}, {}\n{}".format(
            jd['employer'], jd['city'], jd['location2'], jd['title']
        )
        right_side = "start date - end date"
        table.getCellByName("A1").setString(left_side)
        table.getCellByName("B1").setString(right_side)

        # all_tables = current_doc.getTextTables()
        # some_table = all_tables.getByIndex(0)
        # cursor = some_table.createCursorByCellName("A1")
        # for cellname in the_table.getCellNames():
        #     print(cellname)
        #     cell = the_table.getCellByName(cellname)
        #     cell.setString("This cell is {}".format(cellname))
        return table

    def terminate(self):
        fn = time.time()*1000.0
        self.current_writer_doc.storeAsURL("file:///~/temp/{}.odt".format(fn),
                                           [])
        print ("Saved as {}".format(fn))
        # self.current_writer_doc.close(1)

def main():
    creator=ResumeCreator()
    setup = creator.do_set_up()
    new_file = creator.create_new_file()

    creator.insert_job(JOB_DETAILS1)
    creator.terminate()

if __name__ == "__main__":
    main()




# def create_a_table(file, rows, columns):
#     table_string = "com.sun.star.text.TextTable"  # https://wiki.openoffice.org/wiki/Writer/API/Tables
#     the_table = file.createInstance(table_string)
#     the_table.initialize(rows, columns)
#     return the_table
#
#
# def add_item_to_file(item, file):
#     the_text = file.Text
#     the_cursor = the_text.createTextCursor()
#     the_text.insertString(the_cursor, "Behold a table", 0)
#     the_table = create_a_table(file, item, item)
#
#     the_text.insertTextContent(the_cursor, the_table, 0)
#

#
# def add_content_to_file(content, file):
#     for item in content:
#         add_item_to_file(item, file)
#
# def some_test_code(new_file):
#     content = [7, 3, 5]
#     add_content_to_file(content, new_file)
