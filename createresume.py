import uno

JOB_DETAILS1 = {
    'employer': 'University of Michigan Information and Technology Services',
    'city': 'Ann Arbor',
    'state': 'MI',
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
        table_string = "com.sun.star.text.TextTable"  # https://wiki.openoffice.org/wiki/Writer/API/Tables
        the_table = self.current_writer_doc.createInstance(table_string)
        the_table.initialize(rows, columns)
        return the_table

    def create_new_file(self):
        try:
            # https://wiki.openoffice.org/wiki/Writer/API/Overview#Creating.2C_opening_a_Writer_Document
            newdoc = self.desktop.loadComponentFromURL(
                "private:factory/swriter", "_blank", 0, [])
        except:
            raise RuntimeWarning("Creating a new doc didn't work.")
        self.current_writer_doc = newdoc
        return newdoc

    def insert_job(self, job_details, current_doc=None):
        current_doc = current_doc or self.current_writer_doc
        table = self.make_table_and_paragraph_for_job(job_details, current_doc)
        self.insert_table(table)
        self.insert_rest_of_the_junk(job_details)

    def insert_rest_of_the_junk(self, job_details):
        text = self.current_writer_doc.Text
        cursor = text.createTextCursor()
        text.insertString(cursor, job_details.__repr__(), 0)

    def insert_table(self, table):
        text = self.current_writer_doc.Text
        cursor = text.createTextCursor()
        text.insertTextContent(cursor, table, 0)

    def make_table_and_paragraph_for_job(self, job_details, current_doc=None):
        current_doc = current_doc or self.current_writer_doc
        self.current_writer_doc = current_doc
        the_table = self.create_table(1, 2)
        self.insert_table(the_table)

        all_tables = current_doc.getTextTables()
        some_table = all_tables.getByIndex(0)
        cursor = some_table.createCursorByCellName("A1")
        for cellname in the_table.getCellNames():
            print(cellname)
            cell = the_table.getCellByName(cellname)
            cell.setString("This cell is {}".format(cellname))
        return the_table



def main():
    creator=ResumeCreator()
    setup = creator.do_set_up()
    new_file = creator.create_new_file()

    creator.insert_job(JOB_DETAILS1)

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
