<?php

/**
 * Our class for working with Google Sheet API
 *
 * @since 1.0.0
 */
defined( 'ABSPATH' ) || exit;

class NOrderSheet_API {

	private $api_key = null;
	private $service = null;

	/**
	 * NOrderSheet_API constructor.
	 *
	 * @since 1.0.0
	 */
	public function __construct() {

		require NORDER_SHEET_PLUGIN_PATH . '/vendor/autoload.php';

	}

	/**
	 * Returns an authorized API client.
	 *
	 * @since 1.0.0
	 */
	public function getClient() {

		$client = new Google_Client();
		$client->setApplicationName( 'Google Sheets API PHP Quickstart' );
		$client->setScopes( Google_Service_Sheets::DRIVE );
		$client->setAuthConfig( NORDER_SHEET_PLUGIN_PATH . 'credentials.json' );
		$client->setAccessType( 'offline' );
		$client->setPrompt( 'select_account consent' );

		$tokenPath = NORDER_SHEET_PLUGIN_PATH . 'token.json';
		if ( file_exists( $tokenPath ) ) {
			$accessToken = json_decode( file_get_contents( $tokenPath ), true );
			$client->setAccessToken( $accessToken );
		}

		if ( isset( $_POST['norder_auth_code'] ) && ! empty( $_POST['norder_auth_code'] ) ) {
			$authCode    = trim( $_POST['norder_auth_code'] );
			$accessToken = $client->fetchAccessTokenWithAuthCode( $authCode );
			$client->setAccessToken( $accessToken );

			if ( ! file_exists( dirname( $tokenPath ) ) ) {
				mkdir( dirname( $tokenPath ), 0700, true );
			}
			file_put_contents( $tokenPath, json_encode( $client->getAccessToken() ) );

			if ( array_key_exists( 'error', $accessToken ) ) {
				throw new Exception( join( ', ', $accessToken ) );
			}

			$current_url = esc_url( add_query_arg( array() ) );
			?>
            <script language="JavaScript">
                document.location.href = '<?php esc_attr_e( $current_url ); ?>';
            </script>
			<?php
			wp_die();
		}

		if ( $client->isAccessTokenExpired() ) {
			if ( $client->getRefreshToken() ) {
				$client->fetchAccessTokenWithRefreshToken( $client->getRefreshToken() );
			} else {
				$authUrl = $client->createAuthUrl();
				?>
                <p>For work we need get access to your Google Drive account, please, go to
                    <a href="<?php echo $authUrl; ?>" target="_blank">Google auth page</a> for get a code and place it
                    here</p>
                <form action="#" method="post">
                    <input type="text" name="norder_auth_code" value="" class="n-order-settings-input"
                           placeholder="Google code here">
                    <input type="submit" value="Get access" class="button button-primary">
                </form>
				<?php
				return false;
			}
		} else {
			?>
            <h2>You allready set access token, so now you can make a sheets in <a
                        href="<?php echo get_admin_url( null, 'edit.php?post_type=shop_order' ); ?>">Orders
                    page</a></h2>
			<?php
		}

		return $client;
	}

	/**
	 * Add new sheet to google drive and move user to it
	 *
	 * @since 1.0.0
	 *
	 * @param $order_id
	 *
	 * @throws Exception
	 */
	public function create_sheet( $order_id ) {

		$client = $this->getClient();
		if ( ! $client ) {
			wp_die();
		}

		$this->service = new Google_Service_Sheets( $client );

		$spreadsheet = new Google_Service_Sheets_Spreadsheet( array(
			'properties' => array(
				'title' => 'Order-' . $order_id,
			),
		) );

		$spreadsheet = $this->service->spreadsheets->create( $spreadsheet, array(
			'fields' => 'spreadsheetId',
		) );

		return $spreadsheet->spreadsheetId;
	}

	/**
	 * Write all data to order
	 *
	 * @since 1.0.0
	 *
	 * @param $spreadsheet_id
	 * @param $order_id
	 */
	public function put_order_data( $spreadsheet_id, $order_id ) {

		$order          = wc_get_order( $order_id );
		$merge_requests = array();
		$doors          = array( 1, 2, 3 );
		$items_count    = 0;
		$values         = array(
			array(
				'Shop Copy',
				'',
				'Print date: ' . date( 'm/d/y H:i' ),
			),
			array(
				'Order Specifications Sheet',
			),
			array(
				'Customer',
				$order->get_formatted_billing_full_name(),
			),
			array(
				'Phone',
				$order->get_billing_phone(),
			),
			array(
				'Shipping contact',
				$order->get_formatted_shipping_full_name(),
			),
			array(
				'Shipping phone',
				$order->get_meta( 'shipping_phone', true ),
			),
			array(
				'Shipping Address',
				strip_tags( $order->get_formatted_shipping_address() ),
			),
			array(),
			array(
				'Specifications',
			),
		);

		if ( function_exists( 'wcpdf_get_document' ) ) {
			$invoice = wcpdf_get_document( 'partial_payment_invoice', $order );
			if ( $invoice ) {
				$values[0][1][2] = 'Invoice Number: ' . $invoice->get_number();
			}
		}

		$hidden_order_itemmeta = $this->getHiddenFields();
		$merge_requests[]      = $this->getMergeRequest( 0, 2, 0, 2 );
		$merge_requests[]      = $this->getMergeRequest( 8, 9, 0, 2 );

		foreach ( $order->get_items() as $item_id => $item ) {

			$product        = $item->get_product();
			$variation_name = $item->get_meta( 'finish-option', true );

			$values[] = array(
				' ' . $item->get_quantity() . '  ' . $product->get_title(),
				! empty( $variation_name ) ? $variation_name : '',
			);
			$items_count ++;

			if ( $meta_data = $item->get_formatted_meta_data( '' ) ) {
				foreach ( $meta_data as $meta_id => $meta ) {
					if ( in_array( $meta->key, $hidden_order_itemmeta, true ) ) {
						continue;
					}

					$values[] = array(
						'     ' . wp_kses_post( $meta->display_key ),
						trim( strip_tags( wp_kses_post( $meta->display_value ) ) ),
					);
					$items_count ++;
				}
			}
		}

		$body   = new Google_Service_Sheets_ValueRange( [
			'values' => $values
		] );
		$params = [
			'valueInputOption' => 'USER_ENTERED'
		];
		$range  = 'A1:D' . ( $items_count + 12 );

		$result = $this->service->spreadsheets_values->update( $spreadsheet_id, $range,
			$body, $params );

		$this->setColumnWidths( $spreadsheet_id, 300 );
		$this->setFontStyles( $spreadsheet_id, $items_count, count( $doors ) );

		//Merge cells
		$batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest();
		$batchUpdateRequest->setRequests( $merge_requests );
		$response = $this->service->spreadsheets->batchUpdate( $spreadsheet_id, $batchUpdateRequest );

	}

	/**
	 * Get grid range for merge
	 *
	 * @since 1.0.0
	 *
	 * @param $row_start
	 * @param $row_end
	 * @param $col_start
	 * @param $col_end
	 *
	 * @return Google_Service_Sheets_GridRange
	 */
	private function getGridRange( $row_start, $row_end, $col_start, $col_end ) {

		$range = new Google_Service_Sheets_GridRange();
		$range->setStartRowIndex( $row_start );
		$range->setEndRowIndex( $row_end );

		$range->setStartColumnIndex( $col_start );
		$range->setEndColumnIndex( $col_end );

		$range->setSheetId( 0 );

		return $range;
	}

	/**
	 * Get request for merge cells
	 *
	 * @since 1.0.0
	 *
	 * @param $row_start
	 * @param $row_end
	 * @param $col_start
	 * @param $col_end
	 *
	 * @return Google_Service_Sheets_Request
	 */
	private function getMergeRequest( $row_start, $row_end, $col_start, $col_end ) {

		$grid_range = $this->getGridRange( $row_start, $row_end, $col_start, $col_end );

		$request = new Google_Service_Sheets_MergeCellsRequest();
		$request->setMergeType( 'MERGE_ROWS' );
		$request->setRange( $grid_range );

		$merge_request = new Google_Service_Sheets_Request();
		$merge_request->setMergeCells( $request );

		return $merge_request;
	}

	/**
	 * Change column size
	 *
	 * @since 1.0.0
	 *
	 * @param $spreadsheetID
	 * @param int $size
	 */
	public function setColumnWidths( $spreadsheetID, $size = 200 ) {

		$dimensionRange = new Google_Service_Sheets_DimensionRange( [
			'sheetId'    => 0,
			'dimension'  => 'COLUMNS',
			'startIndex' => 0,
			'endIndex'   => 20
		] );

		$dimensionProperties = new Google_Service_Sheets_DimensionProperties( [
			'pixelSize' => $size
		] );

		$requestBody = [
			'requests' => [
				'updateDimensionProperties' => [
					'range'      => $dimensionRange,
					'properties' => $dimensionProperties,
					'fields'     => 'pixelSize'
				]
			]
		];

		$request = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest( $requestBody );
		$result  = $this->service->spreadsheets->batchUpdate( $spreadsheetID, $request );

	}

	private function setFontStyles( $spreadsheet_id, $items_count, $doors_count ) {

		$requests = [
			new Google_Service_Sheets_Request( [
				'repeatCell' => [
					"range" => [
						"sheetId"          => 0,
						"startRowIndex"    => 0,
						"endRowIndex"      => 2,
						"startColumnIndex" => 0,
						"endColumnIndex"   => 2
					],
					"cell"  => [
						"userEnteredFormat" => [
							"horizontalAlignment" => "LEFT",
							"textFormat"          => [
								"bold"     => true,
								"fontSize" => 24,
							]
						]
					],

					"fields" => "UserEnteredFormat(horizontalAlignment,textFormat)"
				]
			] ),
			new Google_Service_Sheets_Request( [
				'repeatCell' => [

					"range" => [
						"sheetId"          => 0,
						"startRowIndex"    => 2,
						"endRowIndex"      => 7,
						"startColumnIndex" => 0,
						"endColumnIndex"   => 1
					],
					"cell"  => [
						"userEnteredFormat" => [
							"horizontalAlignment" => "LEFT",
							"textFormat"          => [
								"bold"     => true,
								"fontSize" => 10,
							]
						]
					],

					"fields" => "UserEnteredFormat(horizontalAlignment,textFormat)"
				]
			] ),
			new Google_Service_Sheets_Request( [
				'repeatCell' => [

					"range" => [
						"sheetId"          => 0,
						"startRowIndex"    => 2,
						"endRowIndex"      => 7,
						"startColumnIndex" => 1,
						"endColumnIndex"   => 2
					],
					"cell"  => [
						"userEnteredFormat" => [
							"horizontalAlignment" => "LEFT",
							"textFormat"          => [
								"bold"     => false,
								"fontSize" => 10,
							]
						]
					],

					"fields" => "UserEnteredFormat(horizontalAlignment,textFormat)"
				]
			] ),
			new Google_Service_Sheets_Request( [
				'repeatCell' => [

					"range" => [
						"sheetId"          => 0,
						"startRowIndex"    => 8,
						"endRowIndex"      => 9,
						"startColumnIndex" => 0,
						"endColumnIndex"   => 1
					],
					"cell"  => [
						"userEnteredFormat" => [
							"horizontalAlignment" => "LEFT",
							"textFormat"          => [
								"bold"     => false,
								"fontSize" => 18,
							]
						]
					],

					"fields" => "UserEnteredFormat(horizontalAlignment,textFormat)"
				]
			] ),
			new Google_Service_Sheets_Request( [
				'updateBorders' => [
					"range" => [
						"sheetId"          => 0,
						"startRowIndex"    => 9,
						"endRowIndex"      => $items_count + 9,
						"startColumnIndex" => 0,
						"endColumnIndex"   => 3
					],

					"top" => [
						"style" => "SOLID",
						"width" => 3,
					],

					"bottom" => [
						"style" => "SOLID",
						"width" => 3,
					],

					"right" => [
						"style" => "SOLID",
						"width" => 3,
					],

					"left" => [
						"style" => "SOLID",
						"width" => 3,
					],
				]
			] ),
			new Google_Service_Sheets_Request( [
				'updateBorders' => [
					"range" => [
						"sheetId"          => 0,
						"startRowIndex"    => $items_count + 9,
						"endRowIndex"      => $doors_count + $items_count + 9,
						"startColumnIndex" => 0,
						"endColumnIndex"   => 3
					],

					"bottom" => [
						"style" => "SOLID",
						"width" => 3,
					],

					"right" => [
						"style" => "SOLID",
						"width" => 3,
					],

					"left" => [
						"style" => "SOLID",
						"width" => 3,
					],
				]
			] ),
		];

		$batchUpdateRequest = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest( [
			'requests' => $requests
		] );

		$this->service->spreadsheets->batchUpdate( $spreadsheet_id, $batchUpdateRequest );

	}

	/**
	 * Get array of hidden fields
	 *
	 * @since 1.0.0
	 *
	 * @return mixed|void
	 */
	private function getHiddenFields() {

		return apply_filters(
			'woocommerce_hidden_order_itemmeta',
			array(
				'_qty',
				'_tax_class',
				'_product_id',
				'_variation_id',
				'_line_subtotal',
				'_line_subtotal_tax',
				'_line_total',
				'_line_tax',
				'method_id',
				'cost',
				'_reduced_stock',
			)
		);
	}

}