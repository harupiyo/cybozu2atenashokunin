#!/usr/bin/perl
use warnings;
no warnings 'uninitialized';
use strict;
use utf8;
use Encode;
use Text::CSV::Encoded;
use List::MoreUtils;
use Getopt::Long;

my $ENCODING = 'cp932';

sub usageAndExit{
	print encode($ENCODING, <<'EOF');
addressperson.csv を準備した上で以下を実行
$ perl cybozu2atenashokunin.pl > atena.csv 2> warn

成果ファイル：atena.csv
警告ファイル：warn
EOF
	exit;
};

my $filename = 'addressperson.csv';
my ($isConfirm, $withFlagFile);
GetOptions(
	"confirm"			=> \$isConfirm,
	"with-flag-file"	=> \$withFlagFile,
	"help"				=> \&usageAndExit);

main:{
	$isConfirm
		? output_for_kowncheck()
		: output_for_ATENA_SHOKUNIN();
}

sub printn($){
	my $str = shift;
	print encode($ENCODING, $str . "\n");
}

sub output_for_kowncheck{
	my $csv = Text::CSV::Encoded->new({ encoding_in => $ENCODING }) or die $!;
	open my $fh, $filename or die $!;

	# サイボウズ側の情報のうち、今回の処理で必要な項目
	my @need = ( "名前（姓）","名前（名）","会社名","役職名","担当者");
	my $first_row = $csv->getline($fh);
	my $_column = build_getColumn_closure( indexFromName(\@need, $first_row) );
	
	# HEADER
	print outputCsv( 'uid', '面識のある人に◯を入力して下さい', @need );
	
	# BODY
	my $id = 2;	# ２行目からがデータ
	while( my $row = $csv->getline($fh) ) {
		print outputCsv( $id++, '', map { $_column->( $row, $_ ) } @need );
	}

	close $fh or die $!;
}

sub output_for_ATENA_SHOKUNIN{
	my $csv = Text::CSV::Encoded->new({ encoding_in => $ENCODING }) or die $!;
	open my $fh, $filename or die $!;
	
	# 必要レコードはどれか
	my $rh_flag = readFlag() if $withFlagFile;
	
	# サイボウズ側の情報のうち、今回の処理で必要な項目
	my @need = ( "名前（姓）", "名前（名）", "よみ（姓）", "よみ（名）", "連名", "自宅〒", "自宅住所１", "自宅住所２", "自宅住所３", "自宅TEL","自宅FAX","自宅携帯", "会社〒", "会社住所１", "会社住所２", "会社住所３", "会社名フリガナ", "会社名", "部署１", "部署２", "役職名", "担当者", "担当部署", "会社コード", "会社名", "部課", "会社よみ", "郵便番号", "住所", "担当者", "担当部署");

	# HEADER: 宛名職人側
	# 宛名職人ではCSV インポート時に、ヘッダの項目名を見て自動割り当てすることができる。そのためのヒントとして出力。
	print outputCsv( qw( 会社〒 会社住所１ 会社住所２ 会社住所３ 会社名 部署１ 部署２ 役職名 氏名 連名) );
	
	# BODY
	my $first_row = $csv->getline($fh);
	my $_column = build_getColumn_closure( indexFromName(\@need, $first_row) );
	my $_compound = build_compoundColumn_closure( $_column );
	
	my $id = 1;
	while( $id++, my $row = $csv->getline($fh) ) {
		
		if($withFlagFile){
			next unless $rh_flag->{ $id }
		}

		my $rh_profile = getProfile( $row, $_column, $_compound );

		# ラベル出力に必要なもののみ出力
		print outputCsv((
			$rh_profile->{zip}, 		# 会社〒
			$rh_profile->{addr1}, 		# 会社住所１
			$rh_profile->{addr2}, 		# 会社住所２
			$rh_profile->{addr3}, 		# 会社住所３
			$rh_profile->{company},		# 会社名
			$rh_profile->{post1},		# 部署１
			$rh_profile->{post2},		# 部署２
			$rh_profile->{office},		# 役職名
			$rh_profile->{name},		# 氏名
			$rh_profile->{joint_name},	# 連名
		));
	}
	
	close $fh or die $!;
}

# サイボウズの情報行から、個人情報を抽出する
sub getProfile{
	my($row, $_column, $_compound) = @_;
	
	my $name = $_column->( $row, "名前（姓）" ) . ' ' . $_column->( $row, "名前（名）" );
	
	# 会社情報の部課が無い時に個人情報側の部署情報を使う
	my($post1,$post2) = ( $_column->( $row, "部課" ), '' );
	unless($post1) {
		$post1 = $_column->( $row, "部署１" );
		$post2 = $_column->( $row, "部署２" );
	}
	
	# 会社名が入っていないと、宛名職人側では自宅住所を読み込もうとして空白になってしまうので、全角空白でダミーの会社名を入れておく
	my $company = $_column->( $row, "会社名" );
	$company = '　' unless $company;

	# 住所は会社情報、個人情報の{会社|自宅} の３種類ある
	# 会社情報 > 個人・会社 > 個人・自宅
	# の優先順位で採用する
	# 判断基準としては、どこに郵便番号があるかどうかとする
	my ( $zip, $addr1, $addr2, $addr3 );
	my $zip1 = $_column->( $row, "郵便番号" ); 
	my $zip2 = $_column->( $row, "会社〒" ); 
	my $zip3 = $_column->( $row, "自宅〒" ); 
	if( $zip1 ne '' ){
		$zip = $zip1;
		$addr1 = $_column->( $row, "住所" );
	}
	elsif( $zip2 ne '' ){
		$zip = $zip2;
		$addr1 = $_column->( $row, "会社住所１" );
		$addr2 = $_column->( $row, "会社住所２" ), 		
		$addr3 = $_column->( $row, "会社住所３" );
	}
	else{
		$zip = $zip3;
		$addr1 = $_column->( $row, "自宅住所１" );
		$addr2 = $_column->( $row, "自宅住所２" ), 		
		$addr3 = $_column->( $row, "自宅住所３" ), 		
	}
	
	warnNeed( '名前', $name, '"名無し"' );
	warnNeed( '郵便番号', $zip, $name );
	warnNeed( '住所/会社住所1/自宅住所1', $addr1, $name );
	warnLength( '会社名', $company, $name, 60 );
	warnLength( '部課/部署1', $post1, $name, 60 );
	warnLength( '住所/会社住所1/自宅住所1', $addr1, $name, 60 );
	warnLength( '会社住所2/自宅住所2', $addr2, $name, 60 );
	warnLength( '会社住所3/自宅住所3', $addr3, $name, 60 );
	
	return {
		'zip'		=> $zip,
		'addr1'		=> $addr1,
		'addr2'		=> $addr2,
		'addr3'		=> $addr3,
		'company'	=> $company,
		'post1'		=> $post1,
		'post2'		=> $post2,
		'name'		=> $name,
		'office'	=> $_column->( $row, "役職名" ),
		'joint_name'=> $_column->( $row, "連名" ),
	};
}

# CSV 一行目の項目定義から、必要な項目名のインデックス番号を取得する
sub indexFromName{
	my ($ra_need, $row_names) = @_;
	my %index;
	foreach my $keyword (@$ra_need){
		$index{ $keyword } = List::MoreUtils::first_index { $_ eq $keyword } @$row_names;
	}
	return \%index;
}

# flag.csv を用意しておく必要がある
# flag.csv のフォーマット
# 	行番号, フラグ（必要な場合は空白ではない)
sub readFlag{
	my $csv = Text::CSV::Encoded->new({ encoding_in => $ENCODING }) or die $!;
	open my $fh, "flag.csv" or die $!;
	my %flag;
	while( my $row = $csv->getline($fh) ) {
		$flag{ $row->[0] } = 1 if $row->[1] ne '';
	}
	close $fh or die $!;
	return \%flag;
}

sub build_getColumn_closure{
	my $index = shift;
	return sub {
		my($ra_row, $keyword) = @_;
		if( $keyword eq '会社名' ){
			# 会社名は同じ列名のフィールドがあるので、入っている方にする
			# CSV 中の後者(29) の会社名が、会社マスタの会社名を指すのでより正式である
			my $c = snip( $ra_row->[ 29 ] || $ra_row->[ $index->{$keyword} ] );
			if( defined $c ){
				return '' if($c eq '個人');
				return $c;
			}
			return '';
		}else{
			return snip( $ra_row->[ $index->{$keyword} ] );
		}
	};
}

sub build_compoundColumn_closure{
	my $_column = shift;
	# 複数の列を合体させ一つの値として返す
	# 当初は合体させていたが、会社と個人の両方に情報が入っている人が居たので、
	# 最初に情報があったもののみを採用するように変更
	return sub{
		my ($ra_row, $ra_column) = @_;
		my @ret;
		foreach my $keyword (@$ra_column){
			my $str = $_column->( $ra_row, $keyword );
			return $str if defined $str;
		}
		return '';
	};
}

# 前後の空白を取る
sub snip{
	my $str = shift;
	$str =~ s/^\s*(\S+)\s*$/$1/ if defined $str;
	return $str;
}

sub outputCsv{
	my $c = join ',', map {
		my $str = $_;
		$str =~ s/,/，/g;	# CSV なので, を全角カンマにエスケープ
							# なお宛名職人側に取り込む時に、\, や
							# 項目を"" で括る方法では上手く取り込めなかった
		$str ? encode($ENCODING, $str ) : ''
	} @_;
	
	return $c . "\n";
}

sub warnNeed{
	my( $column, $value, $owner ) = @_;
	warn encode($ENCODING, "[警告] ${owner}様の'${column}' が入力されていません。\n")
		unless defined $value;
}

sub warnLength{
	my( $column, $value, $owner, $byte_length ) = @_;
	return unless defined $value;
	warn encode($ENCODING, "[警告] ${owner}様の'${column}'(${value}) は宛名職人では${byte_length}bytes にカットされます。\n")
		if( length(encode($ENCODING, $value)) > $byte_length );
}

