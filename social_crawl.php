<html xml:lang-"en" lang="en">
	<head><title>Crawl Social</title></head>
	<body>
	<?php
    // This php program takes a list of urls, and searches for social media links (facebook, twitter, linkedin, instagram, snapchat, pinterest).
		// Infinite execution time
		ini_set('max_execution_time', 0);

		// Open file
		$file = 'urls.txt';
		$fileTarget = 'urlsdata.txt';
		$fp = fopen($file, 'r') or die('ERROR: Cannot find file');

		if ($fp) {
			$urls = explode("\r", fread($fp, filesize($file)));
		}

		$total = count($urls);

		$chars = array("\n", "\r", "\t", "\0", "\x0B");
		$endChars = array('"', ",", " ");
		$web = array("https://", "http://", "www.");
		$socialMedia = array("facebook.com/", "twitter.com/", "linkedin.com/", "instagram.com/", "snapchat.com/", "pinterest.com/");
		$total2 = count($socialMedia);
		$domainList = array();

		// Loop

		for ($x = 0; $x < $total; $x++){

			// Get URL
			$urlClean = strtolower(str_replace($chars, '', $urls[$x]));

			if ($urlClean == "n/a"){
				$time = 8;
				file_put_contents($fileTarget, "N/A" . "\t" . "Skipped", FILE_APPEND);
			} else {
				// Clean URL further
				if (substr($urlClean, 0, 4) != "http"){
					$urlClean = "http://" . $urlClean;
				}

				// Get Domain
				$domain = str_replace($web, "", $urlClean);
				if (strpos($domain, "/") == true){
					$domain = substr($domain, 0, strpos($domain, "/"));
				}

				// Calculate time
				if (in_array($domain, $domainList) == true){
					$time = 5 - array_search($domain, $domainList);
				} else {
					$time = 0;
				}

				array_unshift($domainList, $domain);
				if (count($domainList) > 4){
					array_pop($domainList);
				}

				// Crawl
				file_put_contents($fileTarget, "$urlClean" . "\t", FILE_APPEND);
				$content = @file_get_contents($urlClean, 'r');

				if ($content){
					// Look for social media
					file_put_contents($fileTarget, "Okay" . "\t", FILE_APPEND);

					$foundCount = 0;
					$data2 = null;
					for ($y = 0; $y < $total2; $y++){
						$social = $socialMedia[$y];
						unset($data);

						if (strpos($content, $social) == true){
							$chunk = '';
							$startBitPosition = strpos($content, $social);
							$endBitPosition = $startBitPosition + 1;

							unset($midBit);
							do {
								$midBit = substr($content, $endBitPosition, 1);
								$endBitPosition++;
							} while (!in_array($midBit, $endChars));

							$data = substr($content, $startBitPosition, $endBitPosition - $startBitPosition - 1);
							$data = html_entity_decode($data);
							$data = str_replace('"', '', $data);
							$data = str_replace('\\/', '/', $data);
							$data = strtolower(str_replace($chars, '', $data));
							$foundCount++;
						} else {
							$data = 'N/A';
						}

						$data2 = $data2 . "\t" . $data;
					}

					// Write file ---------------------------------------------
					file_put_contents($fileTarget, $foundCount . $data2 . "\t", FILE_APPEND);

				} else {
					file_put_contents($fileTarget, "Failed", FILE_APPEND);
				}

			}
			file_put_contents($fileTarget, "\n", FILE_APPEND);

			// Counter ------------------------------------------------

			$p = round(($x + 1) * 100/$total, 2);
			$x2 = $x + 1;
			echo "$x2 of $total - $p% [";
			ob_flush();
        	flush();

			for ($w = 1; $w < $time + 1; $w++){
				echo "$w-";
				sleep(1);
				ob_flush();
				flush();
			}
			echo "]<br />";

			ob_flush();
        	flush();
		};

		// Close file
		fclose($fp);

		echo "<br /><br />--------<br />Done!";
	?>
	</body>
</html>
