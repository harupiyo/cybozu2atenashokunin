#!/usr/bin/perl
use warnings;
use strict;
use utf8;
use Text::CSV_XS;
use Encode;
use List::MoreUtils;
use Getopt::Long;

sub usage{
	print encode('utf8', <<'EOF');
EOF
};

my $filename = 'addressperson.csv';
output_for_ATENA_SHOKUNIN();

sub output_for_ATENA_SHOKUNIN{
	my $csv = Text::CSV_XS->new({ binary => 1}) or die Text::CSV_XS->error_diag();
	open my $fh, "<:encoding(utf8)", $filename or die $!;
	
	# サイボウズ側の情報のうち、今回の処理で必要な項目
	my @need = ( "名前（姓）", "名前（名）", "よみ姓）", "よみ（名）", "連名", "自宅〒", "自宅住所１", "自宅住所２", "自宅住所３", "自宅TEL","自宅FAX","自宅携帯", "会社〒", "会社住所１", "会社住所２", "会社住所３", "会社名フリガナ", "会社名", "部署１", "部署２", "役職名", "担当者", "担当部署", "会社コード", "会社名", "部課", "会社よみ", "郵便番号", "住所", "担当者", "担当部署");

	# HEADER
	my $first_row = $csv->getline($fh);
	print outputCsv( @need );
	
	# BODY
	my $_column = build_getColumn_closure( indexFromName(\@need, $first_row) );
	my $_compound = build_compoundColumn_closure( $_column );
	
	my $id = 1;
	while( $id++, my $row = $csv->getline($fh) ) {

		# 複合列, 複数の列を合体させた一つの値を返す
		# 宛名職人フォーマットに合わせエクスポート
		my @c;
		push @c, $_column->( $row, "会社名" );
		push @c, $_compound->( $row, ["部署１", "部課"] );					# 部署１
		push @c, $_column->( $row, "部署２" );
		push @c, $_column->( $row, "役職名" );
		push @c, $_compound->( $row, ["名前（姓）","名前（名）"] );			# 氏名
		push @c, $_compound->( $row, ["自宅〒","会社〒","郵便番号"] ); 		# 会社〒
		push @c, $_compound->( $row, ["自宅住所１","会社住所１","住所"] ); 	# 会社住所１
		push @c, '';	# フリガナ
		push @c, $_compound->( $row, ["名前（姓）","名前（名）"] );			# 氏名
		push @c, $_column->( $row, "連名" );
		push @c, '';	# グループ
		push @c, '';	# 自宅〒
		push @c, '';	# 自宅住所１
		push @c, '';	# 自宅住所２
		push @c, '';	# 自宅住所３
		push @c, '';	# 自宅TEL
		push @c, '';	# 自宅FAX
		push @c, '';	# 自宅携帯
		push @c, '';	# 自宅PHS
		push @c, '';	# 自宅ﾎﾟｹﾍﾞﾙ
		push @c, '';	# 自宅ID
		push @c, '';	# 自宅E-Mail
		push @c, '';	# 自宅URL
		push @c, $_compound->( $row, ["自宅〒","会社〒","郵便番号"] ); 		# 会社〒
		push @c, $_compound->( $row, ["自宅住所１","会社住所１","住所"] ); 	# 会社住所１
		push @c, $_compound->( $row, ["自宅住所２","会社住所２"] ); 			# 会社住所２
		push @c, $_compound->( $row, ["自宅住所３","会社住所３"] ); 			# 会社住所３
		push @c, '';	# 会社TEL
		push @c, '';	# 会社FAX
		push @c, '';	# 会社携帯
		push @c, '';	# 会社PHS
		push @c, '';	# 会社ﾎﾟｹﾍﾞﾙ
		push @c, '';	# 会社ID
		push @c, '';	# 会社E-Mail
		push @c, '';	# 会社URL
		push @c, '';	# その他〒
		push @c, '';	# その他住所１
		push @c, '';	# その他住所２
		push @c, '';	# その他住所３
		push @c, '';	# その他TEL
		push @c, '';	# その他FAX
		push @c, '';	# その他携帯
		push @c, '';	# その他PHS
		push @c, '';	# その他ﾎﾟｹﾍﾞﾙ
		push @c, '';	# その他ID
		push @c, '';	# その他E-Mail
		push @c, '';	# その他URL
		push @c, '';	# 会社名フリガナ
		push @c, $_column->( $row, "会社名" );	# TODO 上にも同様項目がある。整理したい。
		push @c, $_column->( $row, "部署１" );
		push @c, $_column->( $row, "部署２" );
		push @c, $_column->( $row, "役職名" );

		print outputCsv( @c );
	}
	
	close $fh or die $!;
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

sub build_getColumn_closure{
	my $index = shift;
	return sub {
		my($ra_row, $keyword) = @_;
		# 会社名は同じ列名のフィールドがあるので、両方を合わせたものとする
		return ($keyword eq '会社名')
			? snip( $ra_row->[ $index->{$keyword} ] . $ra_row->[29] )
			: snip( $ra_row->[ $index->{$keyword} ] );
	};
}

sub build_compoundColumn_closure{
	my $_column = shift;
	return sub{
		my ($ra_row, $ra_column) = @_;
		my @ret;
		foreach my $keyword (@$ra_column){
			my $str = $_column->( $ra_row, $keyword );
			push @ret, $str if $str ne '';
		}
		return snip( join ' ', @ret );
	};
}

# 前後の空白を取る
sub snip{
	my $str = shift;
	$str =~ s/^\s*(\S+)\s*$/$1/;
	return $str;
}

sub outputCsv{
	my @csv = @_;
	my $c = join ',', map { encode('utf8', qq("$_")) } @csv;
	return $c . "\n";
}

