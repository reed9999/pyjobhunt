import uno

def set_up():
    global smgr
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
                    "com.sun.star.bridge.UnoUrlResolver", localContext )
    ctx = connect_to_running_office(resolver)
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
    model = get_current_writer_doc(desktop)
    return {
        'resolver': resolver,
        'context': ctx,
        'desktop': desktop,
        'model': model,
    }

def connect_to_running_office(the_resolver):
    return the_resolver.resolve(
        "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")



def get_current_writer_doc(desktop):
    return desktop.getCurrentComponent()



def create_new_file(desktop):
    try:
        #https://wiki.openoffice.org/wiki/Writer/API/Overview#Creating.2C_opening_a_Writer_Document
        newdoc = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0,
                                     [])
    except:
        raise RuntimeWarning("Creating a new doc didn't work.")
    return newdoc

def create_a_table(file, rows, columns):
    table_string = "com.sun.star.text.TextTable"  # https://wiki.openoffice.org/wiki/Writer/API/Tables
    the_table = file.createInstance(table_string)
    the_table.initialize(rows, columns)
    return the_table


def add_item_to_file(item, file):
    the_text = file.Text
    the_cursor = the_text.createTextCursor()
    the_text.insertString(the_cursor, "Behold a table", 0)
    the_table = create_a_table(file, item, item)

    the_text.insertTextContent(the_cursor, the_table, 0)


def add_content_to_file(content, file):
    for item in content:
        add_item_to_file(item, file)

def part1():

    # access the document's text property
    text = model.Text
    cursor = text.createTextCursor()

    text.insertString( cursor, "Bambambooyah", 0 )
    another_cursor = model.getCurrentController().getViewCursor()
    text.insertString( another_cursor, "Howdy Woild", 0 )


def main():
    setup = set_up()
    new_file = create_new_file(setup['desktop'])
    content = [7, 3, 5]
    add_content_to_file(content, new_file)

if __name__ == "__main__":
    main()