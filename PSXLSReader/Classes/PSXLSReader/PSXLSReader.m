//
//  PSXLSReader.m
//  PSXLSReader
//
//  Created by Pan on 2018/5/24.
//

#import "PSXLSReader.h"
#import "xls.h"

static const NSInteger CELL_ID_BLANK = 0x201;

@interface PSCell()
@property (nonatomic, assign, readwrite) PSCellContentType type;
@property (nonatomic, assign, readwrite) NSInteger row;
@property (nonatomic, assign, readwrite) NSInteger col;
@property (nonatomic, assign, readwrite) NSString *columnName;            // "A" ... "Z", "AA"..."ZZZ"
@property (nonatomic, strong, readwrite) NSString *str;        // typeof depends on contentsType
@property (nonatomic, strong, readwrite) NSNumber *val;        // typeof depends on contentsType
@end


@interface PSXLSReader ()

@property (nonatomic, assign) xlsWorkBook *workBook;
@property (nonatomic, assign) NSInteger numSheets;
@property (nonatomic, assign) NSInteger activeWorkSheetID;        // keep last one active
@property (nonatomic, assign) xlsWorkSheet *activeWorkSheet;        // keep last one active
@property (nonatomic, assign) xlsSummaryInfo *summary;
@property (nonatomic, assign) BOOL iterating;
@property (nonatomic, assign) NSInteger lastRowIndex;
@property (nonatomic, assign) NSInteger lastColIndex;
@property (nonatomic, assign) NSStringEncoding encoding;

- (void)setWorkBook:(xlsWorkBook *)wb;

- (void)openSheet:(NSInteger)sheetNum;
- (void)formatContent:(PSCell *)content withCell:(xlsCell *)cell;

@end

@implementation PSXLSReader

+ (PSXLSReader *)readerWithPath:(NSString *)filePath
{
    PSXLSReader            *reader;
    xlsWorkBook            *workBook;
    
    // NSLog(@"sizeof FORMULA=%zd LABELSST=%zd", sizeof(FORMULA), sizeof(LABELSST) );
    const char *file = [filePath cStringUsingEncoding:NSUTF8StringEncoding];
    if((workBook = xls_open(file, "UTF-8"))) {
        reader = [PSXLSReader new];
        [reader setWorkBook:workBook];
    }
    return reader;
}

- (instancetype)init
{
    if((self = [super init])) {
        _activeWorkSheetID = NSNotFound;
        _encoding = NSUTF8StringEncoding;
    }
    return self;
}

- (void)dealloc
{
    xls_close_summaryInfo(_summary);
    xls_close_WS(_activeWorkSheet);
    xls_close_WB(_workBook);
}

- (void)setWorkBook:(xlsWorkBook *)wb
{
    _workBook = wb;
    xls_parseWorkBook(_workBook);
    _numSheets = _workBook->sheets.count;
    _summary = xls_summaryInfo(_workBook);
}

- (NSString *)libaryVersion
{
    return [NSString stringWithCString:xls_getVersion() encoding:NSASCIIStringEncoding];
}

// Sheet Information
- (NSInteger)numberOfSheets
{
    return _numSheets;
}

- (NSString *)sheetNameAtIndex:(NSInteger)idx
{
    return idx < _numSheets ? [NSString stringWithCString:(char *)_workBook->sheets.sheet[idx].name encoding:_encoding] : nil;
}

- (NSInteger)rowsForSheetAtIndex:(NSInteger)idx
{
    [self openSheet:idx];
    NSUInteger numRows = _activeWorkSheet->rows.lastrow + 1;
    return idx < _numSheets ? numRows : 0;
}

- (BOOL)isSheetVisibleAtIndex:(NSUInteger)idx
{
    return idx < _numSheets ? (BOOL)_workBook->sheets.sheet[idx].visibility : NO;
}

- (void)openSheet:(NSInteger)sheetNum
{
    if(sheetNum >= _numSheets) {
        _iterating = true;
        _lastColIndex = UINT32_MAX;
        _lastRowIndex = UINT32_MAX;
    } else {
        if(sheetNum != _activeWorkSheetID) {
            _activeWorkSheetID = sheetNum;
            xls_close_WS(_activeWorkSheet);
            _activeWorkSheet = xls_getWorkSheet(_workBook, (int)sheetNum);
            xls_parseWorkSheet(_activeWorkSheet);
        }
    }
}

- (NSInteger)numberOfRowsInSheet:(NSInteger)sheetIndex
{
    [self openSheet:sheetIndex];
    return _activeWorkSheet->rows.lastrow + 1;
}

- (NSInteger)numberOfColsInSheet:(NSInteger)sheetIndex
{
    [self openSheet:sheetIndex];
    return _activeWorkSheet->rows.lastcol + 1;
}

- (PSCell *)cellInWorkSheetIndex:(NSInteger)sheetNum row:(NSInteger)row column:(NSInteger)column
{
    PSCell *content = [PSCell blankCell];
    
    assert(row && column);
    
    [self startIteratorSheetAtIndex:NSNotFound];
    [self openSheet:sheetNum];
    
    row--;
    column--;
    
    NSUInteger numRows = _activeWorkSheet->rows.lastrow + 1;
    NSUInteger numCols = _activeWorkSheet->rows.lastcol + 1;
    
    for (NSUInteger t=0; t<numRows; t++)
    {
        xlsRow *rowP = &_activeWorkSheet->rows.row[t];
        for (NSUInteger tt=0; tt<numCols; tt++)
        {
            xlsCell    *cell = &rowP->cells.cell[tt];
            // NSLog(@"Looking for %d:%d:%d - testing %d:%d Type: 0x%4.4x  [t=%d tt=%d]", sheetNum, row, col, cell->row, cell->col, cell->id, t, tt);
            if(cell->row < row) break;
            if(cell->row > row) return content;
            
            if(cell->id == CELL_ID_BLANK) continue;    // "Blank" filler cell created by libxls
            
            if(cell->col == column) {
                [self formatContent:content withCell:cell];
                return content;
            }
        }
    }
    
    return content;
}

- (PSCell *)cellInWorkSheetIndex:(NSInteger)sheetNum row:(NSInteger)row columnString:(NSString *)column
{
    const char *colStr = [column cStringUsingEncoding:NSUTF8StringEncoding];
    if(strlen(colStr) > 2 || strlen(colStr) == 0) return [PSCell blankCell];
    
    NSInteger col = colStr[0] - 'A';
    if(col < 0 || col >= 26) return [PSCell blankCell];
    char c = colStr[1];
    if(c) {
        col *= 26;
        NSInteger col2 = c - 'A';
        if(col2 < 0 || col2 >= 26) return [PSCell blankCell];
        col += col2;
    }
    col += 1;
    
    return [self cellInWorkSheetIndex:sheetNum row:row column:col];
}

// Iterate through all cells
- (void)startIteratorSheetAtIndex:(NSInteger)sheetIndex
{
    if(sheetIndex != NSNotFound) {
        [self openSheet:sheetIndex];
        _iterating = true;
        _lastColIndex = 0;
        _lastRowIndex = 0;
    } else {
        _iterating = false;
    }
}

- (PSCell *)nextCell
{
    if(!_iterating) {
        return nil;
    }
    
    PSCell *content = [PSCell blankCell];
    
    NSUInteger rowCount = _activeWorkSheet->rows.lastrow + 1;
    NSUInteger columnCount = _activeWorkSheet->rows.lastcol + 1;
    
    if(_lastRowIndex >= rowCount) {
        return content;
    }
    
    for (NSUInteger t = _lastRowIndex; t < rowCount; t++)
    {
        xlsRow *rowP = &_activeWorkSheet->rows.row[t];
        for (NSInteger column = _lastColIndex; column < columnCount; column++)
        {
            xlsCell *cell = &rowP->cells.cell[column];
            
            if(cell->id == CELL_ID_BLANK) {
                continue;
            }
            _lastColIndex = column + 1;
            [self formatContent:content withCell:cell];
            return content;
        }
        _lastRowIndex++;
        _lastColIndex = 0;
    }
    // don't make iterator false - user can keep asking for cells, they all just be blank ones though
    return content;
}

- (void)formatContent:(PSCell *)content withCell:(xlsCell *)cell
{
    NSUInteger col = cell->col;
    
    content.row = cell->row + 1;
    
    {
        content.col = col + 1;
        char colStr[3];
        if(col < 26) {
            colStr[0] = 'A' + (char)col;
            colStr[1] = '\0';
        } else {
            colStr[0] = 'A' + (char)(col/26);
            colStr[1] = 'A' + (char)(col%26);
        }
        colStr[2] = '\0';
        content.columnName = [NSString stringWithFormat:@"%s", colStr];
    }
    
    switch(cell->id) {
        case 0x0006:    //FORMULA
            // test for formula, if
            if(cell->l == 0) {
                content.type = PSCellContentTypeFloat;
                content.val = [NSNumber numberWithDouble:cell->d];
            } else {
                if(!strcmp((char *)cell->str, "bool")) {
                    BOOL b = (BOOL)cell->d;
                    content.type = PSCellContentTypeBool;
                    content.val = [NSNumber numberWithBool:b];
                    content.str = b ? @"YES" : @"NO";
                } else
                    if(!strcmp((char *)cell->str, "error")) {
                        // FIXME: Why do we convert the double cell->d to NSInteger?
                        NSInteger err = (NSInteger)cell->d;
                        content.type = PSCellContentTypeError;
                        content.val = [NSNumber numberWithInteger:err];
                        content.str = [NSString stringWithFormat:@"%ld", (long)err];
                    } else {
                        content.type = PSCellContentTypeString;
                    }
            }
            break;
        case 0x00FD:    //LABELSST
        case 0x0204:    //LABEL
            content.type = PSCellContentTypeString;
            content.val = [NSNumber numberWithLong:cell->l];    // possible numeric conversion done for you
            break;
        case 0x0203:    //NUMBER
        case 0x027E:    //RK
            content.type = PSCellContentTypeFloat;
            content.val = [NSNumber numberWithDouble:cell->d];
            break;
        default:
            content.type = PSCellContentTypeUnknown;
            break;
    }
    
    if(!content.str) {
        content.str = [NSString stringWithCString:(char *)cell->str encoding:NSUTF8StringEncoding];
    }
}

// Summary Information
- (NSString *)appName        { return _summary->appName    ? [NSString stringWithCString:(char *)_summary->appName        encoding:NSUTF8StringEncoding] : @""; }
- (NSString *)author        { return _summary->author    ? [NSString stringWithCString:(char *)_summary->author        encoding:NSUTF8StringEncoding] : @""; }
- (NSString *)category        { return _summary->category    ? [NSString stringWithCString:(char *)_summary->category        encoding:NSUTF8StringEncoding] : @""; }
- (NSString *)comment        { return _summary->comment    ? [NSString stringWithCString:(char *)_summary->comment        encoding:NSUTF8StringEncoding] : @""; }
- (NSString *)company        { return _summary->company    ? [NSString stringWithCString:(char *)_summary->company        encoding:NSUTF8StringEncoding] : @""; }
- (NSString *)keywords        { return _summary->keywords    ? [NSString stringWithCString:(char *)_summary->keywords        encoding:NSUTF8StringEncoding] : @""; }
- (NSString *)lastAuthor    { return _summary->lastAuthor? [NSString stringWithCString:(char *)_summary->lastAuthor    encoding:NSUTF8StringEncoding] : @""; }
- (NSString *)manager        { return _summary->manager    ? [NSString stringWithCString:(char *)_summary->manager        encoding:NSUTF8StringEncoding] : @""; }
- (NSString *)subject        { return _summary->subject    ? [NSString stringWithCString:(char *)_summary->subject        encoding:NSUTF8StringEncoding] : @""; }
- (NSString *)title            { return _summary->title        ? [NSString stringWithCString:(char *)_summary->title        encoding:NSUTF8StringEncoding] : @""; }

@end
