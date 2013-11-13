#!/usr/bin/perl
use perl5i::2;
use SIF::REST;
use XML::Simple;
use Data::Dumper;
use SIF::AU;
use Spreadsheet::ParseExcel;

# PUPOSE: Read a Spreadsheet Table and generate REST Create

my $debug = ($ARGV[$#ARGV] eq 'DEBUG');
if ($debug) { pop @ARGV }

my $filename = shift;
my $sheet = shift;

# ----------------------------------------------------------------------
# Parse Spreadsheet
# ----------------------------------------------------------------------
my $parser   = Spreadsheet::ParseExcel->new();
my $workbook = $parser->parse($filename);
if ( !defined $workbook ) {
	die $parser->error(), ".\n";
}
for my $worksheet ( $workbook->worksheets() ) {
	my ( $row_min, $row_max ) = $worksheet->row_range();
	my ( $col_min, $col_max ) = $worksheet->col_range();

	if ($sheet && ( $sheet ne $worksheet->get_name()) ) {
		#say "Skipping " . $worksheet->get_name();
		next;
	}

	my @fields = map { $worksheet->get_cell(0, $_)->value() } ($col_min .. $col_max);
	#print join(",", @fields); exit 1;

	# Map data
	for my $row ( 1 .. $row_max ) {
		my @vals = map { 
			eval { 
				$worksheet->get_cell($row, $_)->value() 
			} // "" 
		} ($col_min .. $col_max);
		my %data;
		@data{@fields} = @vals;

		my $obj = eval { _obj($worksheet->get_name(), \%data) };
		if ($@) {
			die join("",
				"ERROR $@",
				"Processing $filename: " . $worksheet->get_name() . "\n",
				"Row $row\n",
				Dumper(\%data),
			);
		};
		if ($debug) {
			say $obj->to_xml_string();
		}
		else {
			my $newobj = _create($obj);
			say ref($newobj) . ' ' . $newobj->RefId;
		}
	}
}

exit 0;

# ----------------------------------------------------------------------
# Create the school object
# ----------------------------------------------------------------------
sub _obj {
	my ($name, $data) = @_;
	my $class = "SIF::AU::$name";
	my $obj = $class->new();
	foreach my $key (keys %$data) {
		my $field = $key;
		$field =~ s|^$name/||;
		# say "Setting $field";
		_set($obj, $field, $data->{$key});
	}
	return $obj;
}

# NOTE: Heavily restricted types:
# 	- Only supports single object
sub _set {
	my ($obj, $field, $val) = @_;

	if ($field =~ m|LOOKUP|) {
		# XXX How to lookup a SchoolInfoLocalId to a SchoolInfoRefId
		# say "Skipping lookup field $field";
		return;
	}

	if ($field =~ m|^([^/]+)/(.+)$|) {
		my $name = $1;
		my $rest = $2;

		# say "$name $rest $val";

		# NOTE: Allow a single level of attribute
		my $atts = "";
		if ($name =~ m|(.+)@(.+)|) {
			$name = $1;
			$atts = $2;
		}

		# Create record object (use existing if possible)
		my $newobj = $obj->$name;
		if (! defined($newobj)) {
			my $c = $obj->xml_field_class($name);
			$newobj = $c->new();
			if ($atts =~ m|(.+)=(.+)$|) {
				my $a = "_$1";
				my $v = $2;
				$newobj->$a($v);
			}
			$obj->$name($newobj);
		}

		# Recursive set object data
		_set($newobj, $rest, $val);
	}

	# SET Local field object
	else {
		# say "$field = $val";
		my $atts = "";
		if ($field =~ m|(.+)@(.+)|) {
			$field = $1;
			$atts = $2;
		}

		# Automatically support types where possible...
		eval {
			my $c = $obj->xml_field_class($field);
			my $x = $c->new();
			$x->xml_text_content($val);
			if ($atts =~ m|(.+)=(.+)$|) {
				my $a = "_$1";
				my $v = $2;
				$x->$a($v);
			}
			$obj->$field($x);
		};
		if ($@) {
			# Fall back to single element, no type
			$obj->$field($val);
		}
	}
}

# ======================================================================
# REST CODE

sub _rest {
	our $sifrest;
	if (! $sifrest) {
		$sifrest = SIF::REST->new({
			endpoint => 'http://siftraining.dd.com.au/nswpoc',
		});
		$sifrest->setupRest();
	}
	return $sifrest;
}

# CREATE:
# 	Use object as URL name (plural/singular)
#	Map 
sub _create {
	my ($obj) = @_;
	# TODO support Multiple create
	
	my $class = ref($obj);
	my $name = $class;
	$name =~ s/^SIF::AU:://g;

	# POST / CREATE
	my $xml;
	my $ret = eval {
		$xml = _rest()->post($name . 's', $name, $obj->to_xml_string());
		if ($xml =~ /$name/) {
			return $class->from_xml($xml);
		}
		else {
			die "No valid XML return\n";
		}
	};
	if ($@) {
		die "ERROR $@.\nINPUT XML = " . $obj->to_xml_string() . "\nOriginal XML = $xml\n";
	}
	return $ret;
}

