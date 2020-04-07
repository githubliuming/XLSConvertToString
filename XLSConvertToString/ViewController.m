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
    self.ecxelPathField.enabled = NO;
    self.stringPathField.enabled = NO;
    // Do any additional setup after loading the view.
}

- (IBAction)excelPathSeleted:(NSButton *)sender {
    
    NSOpenPanel *panel = [NSOpenPanel openPanel];
    [panel setCanChooseFiles:YES];//是否能选择文件file
    [panel setCanChooseDirectories:YES];//是否能打开文件夹
    [panel setAllowsMultipleSelection:NO];//是否允许多选file
    [panel  setAllowedFileTypes:@[@"xls"]];
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
     if(self.ecxelPathField.stringValue.length <= 0)
     {
         [self alert:@"请选择 xls 文件"];
         return;
     }
    if (self.stringPathField.stringValue.length <= 0) {
        [self alert:@"请选择 导出路径"];
        return;
    }
    NSLog(@"excel 路径 = %@",self.ecxelPathField.stringValue);
    DHxlsReader * reader = [DHxlsReader xlsReaderWithPath:self.ecxelPathField.stringValue];
    if (!reader)
    {
        [self alert:@"文件格式不正确!!!"];
        return ;
    }
    
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
                NSString * value = cell.str;
                value = [value stringByReplacingOccurrencesOfString:@"%s" withString:@"%@"];
                value = [value stringByReplacingOccurrencesOfString:@"\"" withString:@"\\\""];
                [str appendFormat:@"\"%@\";",value];
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
                [self alert:[NSString stringWithFormat:@"写入文件失败:%@ \n error = %@",strPath,error.userInfo]];
                NSLog(@"写入文件失败");
                return ;
                
            }
        }
    }
    [[NSWorkspace sharedWorkspace] selectFile:nil inFileViewerRootedAtPath:self.stringPathField.stringValue];
}
- (void)setRepresentedObject:(id)representedObject {
    [super setRepresentedObject:representedObject];

    // Update the view, if already loaded.
}

- (void)alert:(NSString *)msg{
    NSAlert *alert = [[NSAlert alloc] init];
    alert.messageText = @"系统提示:";
    alert.informativeText = msg;
    [alert addButtonWithTitle:@"确定"];
    NSInteger ret = [alert runModal];
    switch(ret){
        default:
            printf("按钮点击了。\n");
            break;
    }
}


@end
