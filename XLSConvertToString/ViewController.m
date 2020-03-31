//
//  ViewController.m
//  XLSConvertToString
//
//  Created by anita on 2020/3/31.
//  Copyright © 2020 anita. All rights reserved.
//

#import "ViewController.h"
#import "DHxlsReader.h"
@interface ViewController()
@property (weak) IBOutlet NSTextField *ecxelPathField;
@property (weak) IBOutlet NSTextField *stringPathField;

@end
@implementation ViewController

- (void)viewDidLoad {
    [super viewDidLoad];

    // Do any additional setup after loading the view.
}

- (IBAction)excelPathSeleted:(NSButton *)sender {
    
    NSOpenPanel *panel = [NSOpenPanel openPanel];
    [panel setCanChooseFiles:YES];//是否能选择文件file
    [panel setCanChooseDirectories:YES];//是否能打开文件夹
    [panel setAllowsMultipleSelection:NO];//是否允许多选file
    [panel  setAllowedFileTypes:@[@"xls",@"xlsx"]];
    NSInteger finded = [panel runModal]; //获取panel的响应
    if (finded == NSModalResponseOK) {
       NSArray * urls = [panel URLs];
        NSString * excelUrl = [[urls firstObject] path];
        self.ecxelPathField.stringValue = excelUrl;
        NSLog(@"excel path = %@",excelUrl);
    }
}

- (IBAction)stringFilePathSeleted:(NSButton *)sender {
    
    NSOpenPanel *panel = [NSOpenPanel openPanel];
    [panel setCanChooseFiles:NO];//是否能选择文件file
    [panel setCanChooseDirectories:YES];//是否能打开文件夹
    [panel setAllowsMultipleSelection:NO];//是否允许多选file
    NSInteger finded = [panel runModal]; //获取panel的响应
    if (finded == NSModalResponseOK)
    {
        NSArray * urls = [panel URLs];
        NSString * excelUrl = [[urls firstObject] path];
        self.stringPathField.stringValue = excelUrl;
        NSLog(@"excel path = %@",excelUrl);
    }
}
- (IBAction)convertBtnCliked:(NSButton *)sender {
    
    NSLog(@"excel 路径 = %@",self.ecxelPathField.stringValue);
    DHxlsReader * reader = [DHxlsReader xlsReaderWithPath:self.ecxelPathField.stringValue];
    
    NSInteger sheetCount = [reader numberOfSheets];
    
    NSInteger row = [reader rowsForSheetAtIndex:0];
    NSInteger cellCount = [reader numberOfColsInSheet:0];
    
    NSMutableArray * strings = [[NSMutableArray alloc] init];
    for (NSInteger i = 0; i < row; i++)
    {
        DHcell * keyCell = [reader cellInWorkSheetIndex:1 row:i + 1 col:1];
        for (NSInteger j = 2; j <=cellCount; j++)
        {
            if (j > strings.count)
            {
                NSMutableString * str = [[NSMutableString alloc] init];
                [strings addObject:str];
            }
            NSMutableString * str = strings[j - 2];
            DHcell * cell = [reader cellInWorkSheetIndex:1 row:i + 1 col: j];
            if (cell.str.length > 0)
            {
                [str appendFormat:@"\n\"%@\" = ",keyCell.str];
                [str appendFormat:@"\"%@\";",cell.str];
            }
            
        }
    }
    NSLog(@"sheetCount = %ld row = %ld cellCount = %ld",(long)sheetCount,row,cellCount);
    
    NSLog(@"str0 = %@",strings);
    [self saveToFiles:strings];
}

- (void)saveToFiles:(NSArray *)strings
{
    for (NSInteger i = 0; i < strings.count; i++)
    {
        NSString * strPath = [self.stringPathField.stringValue stringByAppendingPathComponent:[NSString stringWithFormat:@"%ld.strings",(long)i]];
        NSString * s = strings[i];
        if (s.length > 0)
        {
            NSData * data = [s dataUsingEncoding:NSUTF8StringEncoding];
            NSError * error ;
           BOOL isSucc = [data writeToURL:[NSURL fileURLWithPath:strPath] options:NSDataWritingAtomic error:&error];
            if (error)
            {
                NSLog(@"error = %@",[error userInfo]);
            }
            if (isSucc)
            {
                NSLog(@"写入文件成功");
            }
            else
            {
                NSLog(@"写入文件失败");
            }
        }
    }
}
- (void)setRepresentedObject:(id)representedObject {
    [super setRepresentedObject:representedObject];

    // Update the view, if already loaded.
}


@end
