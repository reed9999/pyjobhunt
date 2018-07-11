#From http://christopher5106.github.io/office/2015/12/06/
# openoffice-libreoffice-automate-your-office-tasks-with-python-macros.html

#See also
#https://askubuntu.com/questions/325163/missing-python-in-libreoffice-organize-macros-menu

#Invoke LibreOffice with
# soffice --writer --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"



import uno

localContext = uno.getComponentContext()
resolver = localContext.ServiceManager.createInstanceWithContext(
				"com.sun.star.bridge.UnoUrlResolver", localContext )

def connect_to_running_office(the_resolver):
    return the_resolver.resolve(
        "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")

ctx = connect_to_running_office(resolver)
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)


def get_current_writer_doc():
    return desktop.getCurrentComponent()

model = get_current_writer_doc()


def part1():

    # access the document's text property
    text = model.Text
    cursor = text.createTextCursor()

    text.insertString( cursor, "Bambambooyah", 0 )
    another_cursor = model.getCurrentController().getViewCursor()
    text.insertString( another_cursor, "Howdy Woild", 0 )


    #https://wiki.openoffice.org/wiki/Writer/API/TextRange --> Paragraphs
    the_enum = model.getText().createEnumeration()
    while the_enum.hasMoreElements():
        elem = the_enum.nextElement()
        is_table = (elem.supportsService("com.sun.star.text.TextTable"))
        print(is_table )
        is_paragraph = (elem.supportsService("com.sun.star.text.Paragraph"))
        print(is_paragraph)

    print(dir(text))



def part2():
    desktop.terminate()

def part3():
    components = desktop.Components
    active_frame = desktop.ActiveFrame
    print(dir(desktop))
    print ("######")
    i = 0
    for comp in (components):
        # print(comp.AllVersions) #UnknownPropertyException
        print(comp.Args)
        print(comp.AutoStyles)
        print(comp.getImplementationName())
        print(comp.getViewData()) #pyuno object (com.sun.star.container.XIndexAccess)0x1ce4d68{implementationName=IndexedPropertyValuesContainer, supportedServices={com.sun.star.document.IndexedPropertyValues}, supportedInterfaces={com.sun.star.container.XIndexContainer,com.sun.star.lang.XServiceInfo,com.sun.star.lang.XTypeProvider,com.sun.star.uno.XWeak}}

        text=comp.Text
        cursor = text.createTextCursor()
        text.insertString( cursor, "THIS IS DOCUMENT {} in part 3".format(i), 1 )

        print("Here are the contents of the doc:\n{}".format(text.String))
        i += 1

    print ("######")
    # print(dir(active_frame))

if True:
    part3()
else:
    part1()
    part2()


#components
"""['AllVersions', 'AllowMacroExecution', 'ApplyFormDesignMode', 'Args', 'AutoStyles', 'AutomaticControlFocus', 'AvailableServiceNames', 'AvailableViewControllerNames', 'BasicLibraries', 'Bookmarks', 'BuildId', 'ChapterNumberingRules', 'CharFontCharSet', 'CharFontCharSetAsian', 'CharFontCharSetComplex', 'CharFontFamily', 'CharFontFamilyAsian', 'CharFontFamilyComplex', 'CharFontName', 'CharFontNameAsian', 'CharFontNameComplex', 'CharFontPitch', 'CharFontPitchAsian', 'CharFontPitchComplex', 'CharFontStyleName', 'CharFontStyleNameAsian', 'CharFontStyleNameComplex', 'CharLocale', 'CharacterCount', 'CmisProperties', 'Controllers', 'CurrentController', 'CurrentSelection', 'DefaultPageMode', 'Delegator', 'DialogLibraries', 'DocumentIndexes', 'DocumentProperties', 'DocumentStorage', 'DocumentSubStoragesNames', 'DrawPage', 'EmbeddedObjects', 'EndnoteSettings', 'Endnotes', 'Events', 'FootnoteSettings', 'Footnotes', 'ForbiddenCharacters', 'GraphicObjects', 'HasValidSignatures', 'HideFieldTips', 'Identifier', 'ImplementationId', 'ImplementationName', 'IndexAutoMarkFileURL', 'InteropGrabBag', 'IsTemplate', 'LibraryContainer', 'LineNumberingProperties', 'Links', 'LocalName', 'Location', 'LockUpdates', 'Modified', 'Namespace', 'NumberFormatSettings', 'NumberFormats', 'NumberingRules', 'PagePrintSettings', 'ParagraphCount', 'Parent', 'Printer', 'PropertySetInfo', 'PropertyToDefault', 'RDFRepository', 'RecordChanges', 'RedlineDisplayType', 'RedlineProtectionKey', 'Redlines', 'ReferenceMarks', 'RuntimeUID', 'ScriptContainer', 'ScriptProvider', 'ShowChanges', 'StringValue', 'StyleFamilies', 'SupportedServiceNames', 'Text', 'TextFieldMasters', 'TextFields', 'TextFrames', 'TextSections', 'TextTables', 'Title', 'TransferDataFlavors', 'TwoDigitYear', 'Types', 'UIConfigurationManager', 'URL', 'UndoManager', 'UntitledPrefix', 'VBAGlobalConstantName', 'ViewData', 'WordCount', 'WordSeparator', 'XForms', 'addCloseListener', 'addContentOrStylesFile', 'addDialog', 'addDocumentEventListener', 'addEventListener', 'addEventListener', 'addEventListener', 'addEventListener', 'addEventListener', 'addMetadataFile', 'addModifyListener', 'addModule', 'addPrintJobListener', 'addPropertyChangeListener', 'addRefreshListener', 'addStorageChangeListener', 'addTitleChangeListener', 'addVetoableChangeListener', 'attachResource', 'canCancelCheckOut', 'canCheckIn', 'canCheckOut', 'cancelCheckOut', 'checkIn', 'checkOut', 'close', 'connectController', 'createClone', 'createDefaultViewController', 'createInstance', 'createInstanceWithArguments', 'createLibrary', 'createReplaceDescriptor', 'createSearchDescriptor', 'createViewController', 'disableSetModified', 'disconnectController', 'dispose', 'disposing', 'enableSetModified', 'findAll', 'findFirst', 'findNext', 'getAllVersions', 'getArgs', 'getAutoStyles', 'getAvailableServiceNames', 'getAvailableViewControllerNames', 'getBookmarks', 'getChapterNumberingRules', 'getControllers', 'getCurrentController', 'getCurrentSelection', 'getDocumentIndexes', 'getDocumentLanguages', 'getDocumentProperties', 'getDocumentStorage', 'getDocumentSubStorage', 'getDocumentSubStoragesNames', 'getDrawPage', 'getElementByMetadataReference', 'getElementByURI', 'getEmbeddedObjects', 'getEndnoteSettings', 'getEndnotes', 'getEvents', 'getFlatParagraphIterator', 'getFootnoteSettings', 'getFootnotes', 'getGraphicObjects', 'getIdentifier', 'getImplementationId', 'getImplementationName', 'getLibraryContainer', 'getLineNumberingProperties', 'getLinks', 'getLocation', 'getMapUnit', 'getMetadataGraphsWithType', 'getNumberFormatSettings', 'getNumberFormats', 'getNumberingRules', 'getPagePrintSettings', 'getParent', 'getPreferredVisualRepresentation', 'getPrinter', 'getPropertyDefault', 'getPropertySetInfo', 'getPropertyState', 'getPropertyStates', 'getPropertyValue', 'getRDFRepository', 'getRedlines', 'getReferenceMarks', 'getRenderer', 'getRendererCount', 'getScriptProvider', 'getSomething', 'getStyleFamilies', 'getSupportedServiceNames', 'getText', 'getTextFieldMasters', 'getTextFields', 'getTextFrames', 'getTextSections', 'getTextTables', 'getTitle', 'getTransferData', 'getTransferDataFlavors', 'getTypes', 'getUIConfigurationManager', 'getURL', 'getUndoManager', 'getUntitledPrefix', 'getViewData', 'getVisualAreaSize', 'getXForms', 'hasControllersLocked', 'hasLocation', 'importMetadataFile', 'initNew', 'isDataFlavorSupported', 'isModified', 'isReadonly', 'isSetModifiedEnabled', 'isVersionable', 'leaseNumber', 'load', 'loadFromStorage', 'loadMetadataFromMedium', 'loadMetadataFromStorage', 'lockControllers', 'notifyDocumentEvent', 'paintTile', 'print', 'printPages', 'queryAdapter', 'queryAggregation', 'queryInterface', 'recoverFromFile', 'reformat', 'refresh', 'releaseNumber', 'releaseNumberForComponent', 'removeCloseListener', 'removeContentOrStylesFile', 'removeDocumentEventListener', 'removeEventListener', 'removeEventListener', 'removeEventListener', 'removeEventListener', 'removeEventListener', 'removeMetadataFile', 'removeModifyListener', 'removePrintJobListener', 'removePropertyChangeListener', 'removeRefreshListener', 'removeStorageChangeListener', 'removeTitleChangeListener', 'removeVetoableChangeListener', 'render', 'replaceAll', 'setCurrentController', 'setDelegator', 'setIdentifier', 'setModified', 'setPagePrintSettings', 'setParent', 'setPrinter', 'setPropertyToDefault', 'setPropertyValue', 'setTitle', 'setViewData', 'setVisualAreaSize', 'store', 'storeAsURL', 'storeMetadataToMedium', 'storeMetadataToStorage', 'storeSelf', 'storeToRecoveryFile', 'storeToStorage', 'storeToURL', 'supportsService', 'switchToStorage', 'unlockControllers', 'updateCmisProperties', 'updateLinks', 'wasModifiedSinceLastSave']
"""

#active_frame
"""
['ActiveFrame', 'ComponentWindow', 'ContainerWindow', 'Controller', 'Creator', 'DispatchRecorderSupplier', 'Frames', 'ImplementationId', 'ImplementationName', 'IndicatorInterception', 'IsHidden', 'LayoutManager', 'Name', 'Properties', 'PropertySetInfo', 'SupportedCommandGroups', 'SupportedServiceNames', 'Title', 'Types', 'UserDefinedAttributes', 'activate', 'addCloseListener', 'addEventListener', 'addFrameActionListener', 'addPropertyChangeListener', 'addTitleChangeListener', 'addVetoableChangeListener', 'close', 'contextChanged', 'createStatusIndicator', 'deactivate', 'dispose', 'disposing', 'findFrame', 'focusGained', 'focusLost', 'getActiveFrame', 'getComponentWindow', 'getConfigurableDispatchInformation', 'getContainerWindow', 'getController', 'getCreator', 'getFrames', 'getImplementationId', 'getImplementationName', 'getName', 'getProperties', 'getPropertyByName', 'getPropertySetInfo', 'getPropertyValue', 'getSupportedCommandGroups', 'getSupportedServiceNames', 'getTitle', 'getTypes', 'hasPropertyByName', 'initialize', 'isActive', 'isTop', 'loadComponentFromURL', 'queryDispatch', 'queryDispatches', 'queryInterface', 'registerDispatchProviderInterceptor', 'releaseDispatchProviderInterceptor', 'removeCloseListener', 'removeEventListener', 'removeFrameActionListener', 'removePropertyChangeListener', 'removeTitleChangeListener', 'removeVetoableChangeListener', 'setActiveFrame', 'setComponent', 'setCreator', 'setName', 'setPropertyValue', 'setTitle', 'supportsService', 'windowActivated', 'windowClosed', 'windowClosing', 'windowDeactivated', 'windowHidden', 'windowMinimized', 'windowMoved', 'windowNormalized', 'windowOpened', 'windowResized', 'windowShown']"""
