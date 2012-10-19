#!/usr/bin/perl -w
use strict;
#use utf8;
#use Encode;
use File::Find;
use Tk;
use Tk::Dialog;
use Spreadsheet::ParseExcel;
my $parse = Spreadsheet::ParseExcel->new;

#
# Global Variable
# 一组多个部件都会使用到的全局变量
my $filename;  # selected file's name
my $dirname;   # selected dir's name
my @allXlsFiles; # all xls file included selecting directory
my $files_frame; # contained select files and dirs
my $files_lbox;	 # display select files and dirs
my $result_frame; # contained search result
my $result_text;  # display search result
my $search_frame; # contained search bar
my $search_entry; # search entry bar
my $search_progress; # display a progress
my $percent_done; # search task compeleted percent
my $search_status_label = ''; # search status


my $mw = MainWindow->new();
$mw->title('XLS-Searcher');
$mw->geometry("700x580");

&init_icon($mw);
&createMainMenu;
&createFiles;
&createSearchBar;
&createDisplyResults;
&bind_events;

MainLoop;

#######################################
# Local API
#######################################
# 创建主菜单栏
sub createMainMenu{

	# general menu config
	my $background = 'white';
	my $relief = 'raised';
	my $font = 'Yahei';

	# create a frame to contain mainmenus
	my $parent = $mw->Frame(
		-background => $background,
		-relief => $relief,
		-bd => 2,
		#-height => 1,
	)->pack(
		-side => 'top',
		-fill => 'x',
	);

	$parent->Menubutton(
		-text => 'File',
		-font => $font,
		-width => 7,
		-tearoff => 0,
		-relief => $relief,
		-background => $background,
		-bd => 0,
		-menuitems => [
			[command => "Select File",
				-command => sub {&selectFile}],	
			[command => "Select Directory",
				-command => sub {&selectDir}],	
			[command => "Exit",
				-command => sub {exit;}],	
		],
	)->pack(
		-side => 'left',
		-padx => 0,
		-anchor => 'nw'
	);

	$parent->Menubutton(
		-text => 'Help',
		-font => $font,
		-width => 7,
		-tearoff => 0,
		-bd => 0,
		-relief => $relief,
		-background => $background,
		-menuitems => [
			[command => "Help",
			-command => sub {&help}],	
			[command => "About XLS-File-Searcher",
			-command => sub {&about}],	
			[command => "Update Content",
			-command => sub {&release}],	
			],
		)->pack(
			-side => 'left',
			-padx => 0,
			-anchor => 'nw'
	);
}

#
# createFiles
# 左侧显示选定的目录或者文件
sub createFiles{

	my $font = 'YaHei';
	$files_frame = $mw->Frame(
		-bd => 2,
		-background => 'white',
	)->pack(
		-side => 'left',
		-anchor => 'nw',
		-fill => 'y',
	);

	$files_lbox = $files_frame->Scrolled("Listbox",
		-scrollbars => 'sw',
		-bd => 0,
		-font => $font,
		-background => 'white',
	)->pack(
		-side => 'top',
		-anchor => 'nw',
		-expand => 1,
		-fill => 'y',
	);
}


#
# createSearchBar
# 搜索条
sub createSearchBar{

	my $font = 'YaHei';
	$search_frame = $mw->Frame(
		-bd => 10,
		-height => 20,
		#-relief => 'sunken',
		-relief => 'raised',
		-background => '#E0FFFF'
	)->pack(
		-side => 'top',
		-anchor => 'nw',
		#-expand => 1,
		-fill => 'x',
	);

	# search bar
	my $s = 'support regular expression';
	$search_entry = $search_frame->Scrolled("Entry",
		-scrollbars => 's',
		-bd => 2,
		-relief => 'groove',
		-font => [-size => 20],
		-background => '#E0FFFF',
		-foreground => '#556B2F',
		-width => 30,
		-textvariable => \$s,
	)->pack(
		-side => 'left',
		-anchor => 'nw',
		-fill => 'y',
		#-expand => 1,
		#-fill => 'both',
	);

	# search button
	$search_frame->Button(
		-text => 'Search',	
		-relief => 'raised',
		-background => '#ADD8E6',
		-font => $font,
		-height => 2,
		-width => 14,
		-command => sub {&search},
	)->pack(
		-side => 'left',
		-anchor => 'nw',
	);

=pod
	# progressbar
	$search_progress = $search_frame->ProgressBar(
		-width => 14,
		-length => 160,
		-from => 0,
		-to => 100,
		-blocks => 100,
		#-colors => [0, 'red', 50, 'yellow', 80, 'green'],
		-variable => \$percent_done,
	)->pack(
		-side => 'left',
		-padx => 3,
		-pady => 6,
		-anchor => 'nw',
	);
=cut
	$search_status_label = $search_frame->Label(
		-width => 40,
		-height => 3,
		-font => $font,
		-textvariable => \$percent_done,
	)->pack(
		-side => 'right',
		-padx => 16,
	);

}

#
# createDisplyResults
# 显示查询结果
sub createDisplyResults{
	
	my $font = 'YaHei';
	$result_frame = $mw->Frame(
		-bd => 2,
		-background => 'grey',
	)->pack(
		-side => 'left',
		-anchor => 'nw',
		-expand => 1,
		-fill => 'both',
	);

	$result_text = $result_frame->Scrolled("Text",
		-scrollbars => 'e',
		-bd => 0,
		-font => $font,
		-state => 'disabled',
		-background => '#98FB98',
	)->pack(
		-side => 'top',
		-anchor => 'nw',
		-expand => 1,
		-fill => 'both',
	);
}


#
# selectFile
# 选择文件
sub selectFile{
	
	$filename = $mw->getOpenFile;

	$files_lbox->delete(0, 'end');
	$files_lbox->insert('end', $filename);
	$files_lbox->configure(-background => '#E0FFFF');

}

#
# selectFile
# 选择目录
sub selectDir{
	
	$dirname = $mw->chooseDirectory(
		-title => '!! Please Select A Including .xls File Directory',
		-mustexist => 'true',
		-initialdir => '/',
	);

	print "$dirname\n\n";
	#$dirname =~ s#\/#\\\\#g;
	#print "$dirname\n\n";
	#find(sub {push (@allXlsFiles, $File::Find::name)}, decode('utf8', $dirname));

	# 先置空
	$filename = '';
	@allXlsFiles = ();
	find(sub {push (@allXlsFiles, grep (/.*\.xls/, $File::Find::name))}, $dirname);
	#@allXlsFiles = grep (/.*\.xls$/, @allXlsFiles);

	$files_lbox->delete(0, 'end');
	$files_lbox->insert('end', @allXlsFiles);
	$percent_done = "Total include " . scalar(@allXlsFiles) . " .xls files";

	$files_lbox->configure(-background => '#E0FFFF');

	#&BindMouseWheel($files_lbox);

}

#
# search
# 查询匹配内容
sub search{
	my $ms = $search_entry->get();
	$result_text->configure(-state => 'normal');

	if ($filename && $filename =~ /.*\.xls/ && $ms) {

		$result_text->delete('1.0', "end");
		my @tmp = ($filename);
		&catchU((\@tmp), $ms);
		$result_text->configure(-state => 'disabled');

	} elsif (length(@allXlsFiles) > 0 && $ms){
			
		$result_text->delete('1.0', "end");
		&catchU(\@allXlsFiles, $ms);
		$result_text->configure(-state => 'disabled');

	}
	$result_text->configure(-state => 'disabled');

}

#
# help
# 帮助信息
sub help{
	my $msg = '
author:pangMei
	';
	my $h = $mw->Dialog(
		-title => 'Help',
		-foreground => 'blue',
		-text => $msg,
   	);
	$h->Show;
}

#
# about
# 软件信息
sub about{
	my $msg = '
	=== XLS-File-Searcher  ===
mail: bloodiron888@gmail.com
	';
	my $h = $mw->Dialog(
		-title => 'About',
		-foreground => 'blue',
		-text => $msg,
   	);
	$h->Show;
}

#
# release
# 更新文档
sub release{
	my $msg = '
....
	';
	my $h = $mw->Dialog(
		-title => 'Debug items',
		-foreground => 'blue',
		-text => $msg,
   	);
	$h->Show;
}

# 滚轮支持
sub BindMouseWheel {

    my($w) = @_;
#	$w->focus;

    if ($^O eq 'MSWin32') {
        $w->bind('<MouseWheel>' =>
            [ sub { $_[0]->yview('scroll', -($_[1] / 120) * 3, 'units') },
                Ev('D') ]
        );
    } else {

       # Support for mousewheels on Linux commonly comes through
       # mapping the wheel to buttons 4 and 5.  If you have a
       # mousewheel ensure that the mouse protocol is set to
       # "IMPS/2" in your /etc/X11/XF86Config (or XF86Config-4)
       # file:
       #
       # Section "InputDevice"
       #     Identifier  "Mouse0"
       #     Driver      "mouse"
       #     Option      "Device" "/dev/mouse"
       #     Option      "Protocol" "IMPS/2"
       #     Option      "Emulate3Buttons" "off"
       #     Option      "ZAxisMapping" "4 5"
       # EndSection

        $w->bind('<4>' => sub {
            $_[0]->yview('scroll', -3, 'units') unless $Tk::strictMotif;
        });

        $w->bind('<5>' => sub {
                  $_[0]->yview('scroll', +3, 'units') unless $Tk::strictMotif;
        });
    }

} # end BindMouseWheel

#
# catchU
# 从指定的目录和文件中查找字符串
sub catchU{
	#$result_text->tagConfigure('filename', -foreground => '#228B22');
	$result_text->tagConfigure('filename', -foreground => '#008000');
	$result_text->tagConfigure('location', -background => '#FFFF00', -foreground => 'black');
	$search_status_label->configure(-background => '#191970', -foreground => '#D87093');
	my $i = 1;
	my $j = 0;
	my $xlses = shift;
	my $str = shift;
	for (@$xlses){
		print "filename $_\n";
		my $workbook = $parse->parse($_);

		if (!defined $workbook){
			die $parse->error . "\n";
		}

		for my $worksheets ($workbook->worksheets){

			my ($row_min, $row_max) = $worksheets->row_range;
			my ($col_min, $col_max) = $worksheets->col_range;

			for my $row ($row_min .. $row_max){
				for my $col ($col_min .. $col_max){
					my $cell = $worksheets->get_cell($row, $col);
					next unless $cell;	
					my $value = $cell->value;
					if ($value =~ /$str/){
						$j++;
						my $nrow = $row + 1;
						my $ncol = $col + 1;
						#$result_text->insert('1.0', " $value\n");
						#$result_text->insert('1.0', "Location:[$nrow,$ncol]", 'location');
						#$result_text->insert('1.0', "$_ ", 'filename');
						$result_text->insert('end', "row-col:[$nrow,$ncol]", 'location');
						$result_text->insert('end', "$_ ", 'filename');
						$result_text->insert('end', " $value\n");
						$result_text->update;
					}
				}
			}

		}
		# 显示搜索进度
		$percent_done = "Searching...\nsearched " . ($i++ + 1). " files" . "\ndiscovered ${j} files include <$str>";
		$search_status_label->update;
	}
		# 显示最后的搜索统计结果
		$search_status_label->configure(-background => '#AFEEEE', -foreground => '#191970');
		$percent_done = "Task compeleted\nFiles sum: " . ($i-1) .  "\nmatches sum: ${j}";
		$search_status_label->update;
}

# 
# text_right_click 
# 文档区域右键显示选项
sub text_right_click{
	my $popmenu = $mw->Menu(
		-menuitems => [
			[
				"command" => "Save Result",
				-command => sub {&save_result},	
			],	
			[
				"command" => "Clean Result",
				-command => sub {&clean_result},	
			],	
		],	
		-tearoff => 0,
		-relief => 'groove',
	);
	$popmenu->Popup(
		# 弹出的位置在鼠标这里
		-popover => 'cursor',
		-popanchor => 'nw',
	);
}

# 
# save_result 
# 保存查询结果
sub save_result{

	my $textContent = $result_text->get('1.0', 'end');
	if ($textContent ne "\n"){
		my $saveFile = $mw->getSaveFile(
			-title => 'Save File',
			-initialdir => '/',
			-defaultextension => '.txt',
		);

		print length ($textContent);
		open my $s, "> $saveFile"
			or die "$!";
		print $s $textContent;
		close $s;
	}
}

#
# clean_result
# 清空结果输出
sub clean_result{

	my $textState = $result_text->configure(-state);
	$result_text->configure(-state => 'normal');
	$result_text->delete('1.0', "end");
	$result_text->configure(-state => 'disabled');

}

#
# bind events
#
sub bind_events{
	$files_lbox->bind(
		'<Button-1>',
		sub{
			$files_lbox->focus;
			&BindMouseWheel($files_lbox);
		}
	);

	$result_text->bind(
		'<Button-1>',
		sub{
			$result_text->focus;
			&BindMouseWheel($result_text);
		}
	);

# disable text widget right click menu
# then bind my defined menu
	$result_text->menu(undef);
	$result_text->bind(
		'<Button-3>',
		sub {
			&text_right_click;
		}
	);
}

#
# init_icon
# 创建预览条图标
sub init_icon{
	my $parent = shift;
	open my $icon, '> Camel.xpm'
		or die '$!';

print $icon <<'EOF';
/* XPM */
static char *Camel[] = {
/* width height num_colors chars_per_pixel */
"    32    32        2            1",
/* colors */
". c #696969",
"# c #000000",
/* pixels */
"................................",
"................................",
"...................###..........",
".......####......######.........",
"....####.##.....########........",
"....########....#########.......",
"......######..###########.......",
"......#####..#############......",
".....######.##############......",
".....######.###############.....",
".....######################.....",
".....#######################....",
".....#######################....",
"......#######################...",
".......####################.#...",
"........###################.#...",
"........###############.###.#...",
"............#######.###.###.#...",
"............###.###.##...##.....",
"............###.###..#...##.....",
"............##.####..#....#.....",
"............##.###...#....#.....",
"............##.##...#.....#.....",
"............#...#...#.....#.....",
"............#....#..#.....#.....",
"............#.....#.#.....#.....",
"............#.....###.....#.....",
"...........##....##.#....#......",
"...........#..............#.....",
".........###.............#......"
"................................",
"................................",
};
EOF

	close $icon;

	my $image = $parent->Photo(-file => 'Camel.xpm');
	$parent->Icon(-image => $image);
}
