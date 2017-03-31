<?php
/**
 * This is a small script that utilizes PHPExcel to parse .xlsx sheets
 * Created specifically for Level time sheets and nothing else. Longbeard is a synonym of Sage.
 *
 * No parameters required. Add your files to the source folder and run in cli or in a browser.
 * Outputs in terminal/cli or html.
 *
 * User name should be in cell A3 of the sheet if per user tally is necessary.
 *
 * @author Chris Rubio <chris.rubio@level-studios.com>
 */

// Setup defaults
error_reporting(0);
ini_set('memory_limit','1600M');
set_include_path(get_include_path() . PATH_SEPARATOR . 'classes/');

include 'PHPExcel/IOFactory.php';

// Measurement
$time_start = microtime(true);

// Directory information
$dir = "./source/";
$all_files = scandir($dir);
$files = array_diff($all_files, array('.', '..'));
$file_count = count($files);

// User rates array from csv
include "./csv/Resource_BillRate.php";

// Running total dataset
$totals = [];

// Valid cells and status
$total_col = "K";
$project_col = "A";

// Project code patterns
$proj_pattern = '/(APP|RMG|TIC)/';
$non_bill_pattern = '/(BDN|OVN)/';

// File parsing
foreach ($files AS $key => $file) {
		
	// Current file to load
	$input_file = $dir.$file;

	// Load excel object
	$excel_obj = PHPExcel_IOFactory::load($input_file);

	// Find the number of rows
	$highest_row = $excel_obj->setActiveSheetIndex(0)->getHighestRow();

	// User name for the worksheet - REQUIRED
	$user = strtolower(get_cell_value($excel_obj, "A", 3));

	// Progress report for CLI
	if (PHP_SAPI === 'cli') {
		if (($key % 2) == 0) {
			echo round(($key/$file_count) * 100)."% completed\r";
		}
	}
	
	// First six rows of data don't need to be parsed
	for ($i=6; $i < $highest_row; $i++) {

		$value = get_cell_value($excel_obj, $total_col, $i);
		
		// Find project code first and match it
		$proj_code = get_cell_value($excel_obj, $project_col, $i);
		
		if (preg_match($proj_pattern, $proj_code) AND $user) {
			$total_value = get_cell_value($excel_obj, $total_col, $i);

			// If its non-billable, it goes under that index
			if (preg_match($non_bill_pattern, $proj_code)) {
				$totals['non-billable'][$proj_code] += $total_value;
				$totals['total-non-billable'] += $total_value;
				$totals['users'][$user]['total-non-billable'] += $total_value;
			}
			else {
				$totals['billable'][$proj_code] += $total_value;
				$totals['total-billable'] += $total_value;
				$totals['users'][$user]['total-billable'] += $total_value;
			}

			// Record all hours for this user
			$totals['users'][$user]['all-hours'] += $total_value;
		}
	}
}

// Calculate revenue per user now that we're done totalling
user_revenue($totals, $rates);

$time_end = microtime(true);
$time = $time_end - $time_start;

//=============
// rendering
//=============
if (PHP_SAPI === 'cli') {
	render_cli($totals, $file_count);

	// Ending remarks
	echo "\nOperation completed in ".round($time,2)." seconds.\n";
	echo "Peak memory usage: ".memory_get_peak_usage();
	echo "\n";
} else {
	render_html($totals, $files, $time);
}

//=============
// helpers
//=============

/**
 * Returns the raw calculated value of one cell
 * @param Object $obj Excel doc object from PHPExcel library 
 * @param String $x cell column
 * @param Integer $y cell row
 * @return String The cell value
 */
function get_cell_value(&$obj, $x = 'A', $y = 1 ) {
	return $obj->getActiveSheet()->getCell($x.$y)->getCalculatedValue();
}

/**
 * Calculates billable hours x employee rate. Returns master array with that data.
 * @param Array $totals The master data array
 * @return Array $totals The master array with calculated index
 */
function user_revenue(&$totals, $rates) {
	// Force lowercase
	$rates = array_change_key_case($rates, CASE_LOWER);
	
	foreach ($totals['users'] AS $username => $hours) {
		// Add total revenue to new key
		$totals['users'][$username]['revenue'] = "$".number_format($rates[$username] * $hours['total-billable'], 2, ".", "," );
	}
}

/**
 * Rendering for the CLI
 * @param  Array $totals The master data array
 */
function render_cli($totals, $file_count) {
	echo "\nTotal project hours by code \n";
	echo "***************************\n";

	echo "\nTotal billable hours: ".$totals['total-billable']."\n";
	echo "Total non-billable hours: ".$totals['total-non-billable']."\n";

	echo "\nFiles parsed: ".$file_count."\n\n";

    // Billable project output
    echo "Project code (billable)\t\tHours\n";
    foreach ($totals['billable'] AS $key => $row) {
    	echo $key."\t\t";
    	echo $row."\t\n";
    }

    // Non billable project output
    echo "\nProject code (non-billable)\tHours\n";
    foreach ($totals['non-billable'] AS $key => $row) {
    	echo $key."\t\t";
    	echo $row."\t\n";
    }

    echo "\nUser name\t\t\tHours(total)\tRevenue(total billable hours)\n";
	foreach ($totals['users'] AS $key => $row) {
		echo $key."\t\t\t";
		echo $row['all-hours']."\t\t";
		echo $row['revenue']."\t\n";
	}
}

/**
 * @param Array $totals The master array
 * @param Integer $file_count Num of files parsed
 * @param Integer $time Total time taken for operation
 */
function render_html($totals, $file_count, $time) {
	echo "<html><head><meta http-equiv='Content-Type' content='text/html'; charset=UTF-8' /></head><body><section><h3>Total project hours by code</h3><span> * Add .xlsx files to sorce folder and refresh page to calculate";

	// Grand total
	echo "<p>Total billable hours: ".$totals['total-billable']."</p>";
	echo "<p>Total <span style='color:red;'>non-billable</span> hours: ".$totals['total-non-billable']."</p>";

	// Billable table
	echo "<table border='1'><thead><tr><th>Project code (billable)</th><th>Total hours</th></tr></thead><tbody>";
    foreach ($totals['billable'] as $key => $row) {
    	echo "<tr>";
    	echo "<td>".$key."</td>";
    	echo "<td>".$row."</td>";
    	echo "</tr>";
    }
    echo "</tbody></table>";

    // Non billable project table
    echo "<table border='1' style='margin-top:20px;'><thead><tr><th>Project code (non-billable)</th><th>Total hours</th></tr></thead><tbody>";
    foreach ($totals['non-billable'] as $key => $row) {
    	echo "<tr>";
    	echo "<td>".$key."</td>";
    	echo "<td>".$row."</td>";
    	echo "</tr>";
    }
    echo "</tbody></table>";

    // By individual
    echo "<table border='1' style='margin-top:20px;'><thead><tr><th>User name</th><th>Total hours</th><th>Billable hours</th><th>Revenue</tr></thead><tbody>";
    foreach ($totals['users'] as $key => $row) {
    	echo "<tr>";
    	echo "<td>".$key."</td>";
    	echo "<td>".$row['all-hours']."</td>";
    	echo "<td>".$row['total-billable']."</td>";
    	echo "<td>".$row['revenue']."</td>";
    	echo "</tr>";
    }
    echo "</tbody></table>";

	// Measurement
	echo "<p>* ".count($file_count)." files parsed in ".round($time, 2)." seconds</p><br />";
    echo "</section></body></html>";
}

?>