//
//  PSCell.m
//  PSXLSReader
//
//  Created by Pan on 2018/5/24.
//

#import "PSCell.h"

@interface PSCell()
@property (nonatomic, assign, readwrite) PSCellContentType type;
@property (nonatomic, assign, readwrite) NSInteger row;
@property (nonatomic, assign, readwrite) NSInteger col;
@property (nonatomic, assign, readwrite) NSString *columnName;            // "A" ... "Z", "AA"..."ZZZ"
@property (nonatomic, strong, readwrite) NSString *str;        // typeof depends on contentsType
@property (nonatomic, strong, readwrite) NSNumber *val;        // typeof depends on contentsType
@end

@implementation PSCell

+ (PSCell *)blankCell
{
    return [[PSCell alloc] init];
}

- (NSString *)description {
    NSMutableString *s = [NSMutableString stringWithCapacity:128];
    
    const char *name;
    switch(self.type) {
        case PSCellContentTypeBlank:        name = "PSCellContentTypeBlank";        break;
        case PSCellContentTypeString:    name = "PSCellContentTypeString";    break;
        case PSCellContentTypeInteger:    name = "PSCellContentTypeInteger";    break;
        case PSCellContentTypeFloat:        name = "PSCellContentTypeFloat";        break;
        case PSCellContentTypeBool:        name = "PSCellContentTypeBool";        break;
        case PSCellContentTypeError:        name = "PSCellContentTypeError";        break;
        default:            name = "PSCellContentTypeUnknown";    break;
    }
    
    [s appendString:@"====================\n"];
    [s appendFormat:@"CellType: %s row=%ld col=%@/%ld\n", name, self.row, self.columnName, self.col];
    [s appendFormat:@"   string:    %@\n", self.str];
    
    switch(self.type) {
        case PSCellContentTypeInteger:    [s appendFormat:@"     long:    %ld\n",    [self.val longValue]];    break;
        case PSCellContentTypeFloat:        [s appendFormat:@"    float:    %lf\n",    [self.val doubleValue]];    break;
        case PSCellContentTypeBool:        [s appendFormat:@"     bool:    %d\n",    [self.val boolValue]];    break;
        case PSCellContentTypeError:        [s appendFormat:@"    error:    %ld\n",    [self.val longValue]];    break;
        default: break;
    }
    return s;
}
@end
