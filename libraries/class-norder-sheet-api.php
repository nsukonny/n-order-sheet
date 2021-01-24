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
	 * Get google client for work with API
	 *
	 * @since 1.0.0
	 *
	 * @return Google_Client
	 * @throws \Google\Exception
	 */
	private function getGoogleClient() {

		$client = new Google_Client();
		$client->setApplicationName( 'Google Sheets API PHP Quickstart' );
		$client->setScopes( Google_Service_Sheets::DRIVE );
		$client->setAuthConfig( NORDER_SHEET_PLUGIN_PATH . 'credentials.json' );
		$client->setAccessType( 'offline' );
		$client->setPrompt( 'select_account consent' );

		if ( file_exists( NORDER_SHEET_TOKEN_PATH ) ) {
			$accessToken = json_decode( file_get_contents( NORDER_SHEET_TOKEN_PATH ), true );
			$client->setAccessToken( $accessToken );
		}

		return $client;
	}

	/**
	 * Delete access token
	 *
	 * @since 1.0.0
	 *
	 * @return bool
	 */
	public static function deleteAccessToken() {

		if ( file_exists( NORDER_SHEET_TOKEN_PATH ) ) {
			return unlink( NORDER_SHEET_TOKEN_PATH );
		}

		return true;
	}

	/**
	 * Returns an authorized API client.
	 *
	 * @since 1.0.0
	 * @throws \Google\Exception
	 */
	public function getClient() {

		$client = $this->getGoogleClient();

		if ( isset( $_POST['norder_auth_code'] ) && ! empty( $_POST['norder_auth_code'] ) ) {
			$authCode    = trim( $_POST['norder_auth_code'] );
			$accessToken = $client->fetchAccessTokenWithAuthCode( $authCode );
			$client->setAccessToken( $accessToken );

			if ( ! file_exists( dirname( NORDER_SHEET_TOKEN_PATH ) ) ) {
				mkdir( dirname( NORDER_SHEET_TOKEN_PATH ), 0700, true );
			}
			file_put_contents( NORDER_SHEET_TOKEN_PATH, json_encode( $client->getAccessToken() ) );

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
				?>
                <h2>You allready set access token, so now you can make a sheets in <a
                            href="<?php echo get_admin_url( null, 'edit.php?post_type=shop_order' ); ?>">Orders
                        page</a></h2>
                <p>
                    <a href="<?php echo get_admin_url( null, 'admin.php?page=n-order-sheet&act=remove_token' ); ?>">Reset
                        access token</a> |
                    <a href="<?php echo get_admin_url( null, 'admin.php?page=n-order-sheet&act=clear_sheets' ); ?>">Clear
                        all sheets</a>
                </p>
				<?php
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
            <p>
                <a href="<?php echo get_admin_url( null, 'admin.php?page=n-order-sheet&act=remove_token' ); ?>">Reset
                    access token</a> |
                <a href="<?php echo get_admin_url( null, 'admin.php?page=n-order-sheet&act=clear_sheets' ); ?>">Clear
                    all sheets</a>
            </p>
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

		$order = wc_get_order( $order_id );

		$merge_requests = array();
		$doors          = array( 1, 2, 3 );
		$items_count    = 0;

		$client                      = array();
		$client['full_billing_name'] = $order->get_formatted_billing_full_name();
		if ( empty( $client['full_billing_name'] ) ) {
			$client['full_billing_name'] = $order->get_formatted_shipping_full_name();
		}

		$client['full_shipping_name'] = $order->get_formatted_shipping_full_name();
		if ( empty( $client['full_shipping_name'] ) ) {
			$client['full_shipping_name'] = $order->get_formatted_billing_full_name();
		}

		$client['shipping_address']        = ! empty( $order->get_shipping_country() ) ? $order->get_shipping_country() : '';
		$client['shipping_address']        .= ! empty( $order->get_shipping_city() ) ? ', ' . $order->get_shipping_city() : '';
		$client['shipping_address']        .= ! empty( $order->get_shipping_state() ) ? ', ' . $order->get_shipping_state() : '';
		$client['shipping_address']        .= ! empty( $order->get_shipping_postcode() ) ? ', ' . $order->get_shipping_postcode() : '';
		$client['full_shipping_address_1'] = $order->get_shipping_address_1() . '<br>';

		if ( empty( $order->get_shipping_address_1() ) ) {
			$client['shipping_address']        = ! empty( $order->get_billing_country() ) ? $order->get_billing_country() : '';
			$client['shipping_address']        .= ! empty( $order->get_billing_city() ) ? ', ' . $order->get_billing_city() : '';
			$client['shipping_address']        .= ! empty( $order->get_billing_state() ) ? ', ' . $order->get_billing_state() : '';
			$client['shipping_address']        .= ! empty( $order->get_billing_postcode() ) ? ', ' . $order->get_billing_postcode() : '';
			$client['full_shipping_address_1'] = $order->get_billing_address_1();
		}

		$client['full_shipping_address_2'] = $order->get_shipping_address_2();
		if ( empty( $order->get_shipping_address_2() ) ) {
			$client['full_shipping_address_2'] = $order->get_billing_address_2();
		}

		$invoice_number = '';
		if ( function_exists( 'wcpdf_get_document' ) ) {
			$invoice = wcpdf_get_document( 'invoice', $order->get_id() );
			if ( $invoice ) {
				$invoice_number = $invoice->get_number()->formatted_number;
			}
		} else {
			$invoice_number = get_post_meta( $order->get_id(), '_wcpdf_invoice_number', true );
		}

		$scheduled_ship_date = '';
		$values              = array(
			array(
				'',
				'Shop Copy',
				'',
				'Order #:',
				$order_id,
			),
			array(
				'',
				'Order Specifications Sheet',
				'',
				'Print Date:',
				date( 'm/d/y H:i' )
			),
			array(
				'',
				'Customer',
				$client['full_billing_name'],
				'Invoice Number:',
				$invoice_number,
			),
			array(
				'',
				'Phone',
				$order->get_billing_phone(),
				'Scheduled Ship Date:',
				$scheduled_ship_date,
			),
			array(),
			array(
				'',
				'Shipping contact',
				$client['full_shipping_name'],
			),
			array(
				'',
				'Shipping phone',
				$order->get_meta( 'shipping_phone', true ),
			),
			array(
				'',
				'Shipping Address',
				strip_tags( $client['full_shipping_address_1'] ),
			),
			array(
				'',
				'',
				strip_tags( $client['shipping_address'] ),
			),
		);

		if ( ! empty( $client['full_shipping_address_2'] ) ) {
			$values[] = array(
				'',
				'Shipping Address 2',
				strip_tags( $client['full_shipping_address_2'] ),
			);

			$values[] = array(
				'',
				'',
				strip_tags( $client['shipping_address'] ),
			);
		}

		$customer_lines_count = count( $values );

		$values[] = array();
		$values[] = array(
			'',
			'Specifications',
		);

		$hidden_order_itemmeta = $this->getHiddenFields();
		$merge_requests[]      = $this->getMergeRequest( 0, 2, 1, 3 );
		$merge_requests[]      = $this->getMergeRequest( $customer_lines_count, $customer_lines_count + 1, 1, 3 );
		foreach ( $order->get_items() as $item_id => $item ) {

			$product        = $item->get_product();
			$variation_name = $item->get_meta( 'finish-option', true );
			$meta_data      = $item->get_formatted_meta_data( '' );
			if ( $meta_data ) {

				foreach ( $meta_data as $meta_id => $meta ) {
					if ( in_array( $meta->key, $hidden_order_itemmeta, true ) || '_' === $meta->display_key[0] ) {
						unset( $meta_data[ $meta_id ] );
					}
				}

				if ( 3 >= count( $meta_data ) ) {
					foreach ( $meta_data as $meta ) {
						$variation_name .= ' ' . trim( strip_tags( wp_kses_post( $meta->display_value ) ) );
					}
				}
			}

			$values[] = array(
				'',
				' ' . $item->get_quantity() . ' x ' . $product->get_title(),
				! empty( $variation_name ) ? $variation_name : '',
			);
			$items_count ++;

			if ( $meta_data ) {

				if ( 3 < count( $meta_data ) ) {
					$attributes_showed = false;
					$attributes        = $this->prepare_reorganized_attributes( $meta_data );

					foreach ( $meta_data as $meta_id => $meta ) {

						if ( isset( $attributes[ $meta->key ] ) ) {
							if ( ! $attributes_showed ) {
								foreach ( $attributes as $attribute ) {
									if ( empty( $attribute ) ) {
										continue;
									}

									if ( '_uni_cpo_rush_option' == $attribute->key || '_uni_cpo_door_thickness' == $attribute->key ) {
										$values[] = array();
										$items_count ++;
									}

									$values[] = array(
										'',
										'     ' . wp_kses_post( $attribute->display_key ),
										trim( strip_tags( wp_kses_post( $attribute->display_value ) ) ),
									);
									$items_count ++;

									if ( '_uni_cpo_door_thickness' == $attribute->key ) {
										$values[]    = array();
										$values[]    = array();
										$items_count += 2;
									}
								}

								$attributes_showed = true;
							}

							continue;
						}

						if ( '_uni_cpo_rush_option' == $meta->key || '_uni_cpo_door_thickness' == $meta->key ) {
							$values[] = array();
							$items_count ++;
						}

						$values[] = array(
							'',
							'     ' . wp_kses_post( $meta->display_key ),
							trim( strip_tags( wp_kses_post( $meta->display_value ) ) ),
						);
						$items_count ++;

						if ( '_uni_cpo_door_thickness' == $meta->key ) {
							$values[]    = array();
							$values[]    = array();
							$items_count += 2;
						}
					}
				}

			}
		}

		$total_row               = $items_count + $customer_lines_count + 2;
		$values[ $total_row ]    = array( '', ' DOORS SUMMARY' );
		$values[ ++ $total_row ] = array( '', ' Door Unit Total Dimensions:' );
		$values[ ++ $total_row ] = array( '', ' Suggested Rough Opening:' );

		$body   = new Google_Service_Sheets_ValueRange( [
			'values' => $values
		] );
		$params = [
			'valueInputOption' => 'USER_ENTERED'
		];
		$range  = 'A1:F' . ( $items_count + $customer_lines_count + 5 );

		$result = $this->service->spreadsheets_values->update( $spreadsheet_id, $range,
			$body, $params );

		$this->setSize( $spreadsheet_id, 'COLUMNS', 1, 5, 250 );
		$this->setSize( $spreadsheet_id, 'ROWS', $customer_lines_count, $customer_lines_count + 1, 50 );
		$this->setFontStyles( $spreadsheet_id, $items_count, $customer_lines_count );

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
	 * @param $spreadsheet_id
	 * @param $start_index
	 * @param $end_index
	 * @param string $type
	 * @param int $size
	 */
	public function setSize( $spreadsheet_id, $type, $start_index, $end_index, $size = 200 ) {

		$dimension_range = new Google_Service_Sheets_DimensionRange( [
			'sheetId'    => 0,
			'dimension'  => $type,
			'startIndex' => $start_index,
			'endIndex'   => $end_index,
		] );

		$dimension_properties = new Google_Service_Sheets_DimensionProperties( [
			'pixelSize' => $size
		] );

		$request_body = [
			'requests' => [
				'updateDimensionProperties' => [
					'range'      => $dimension_range,
					'properties' => $dimension_properties,
					'fields'     => 'pixelSize'
				]
			]
		];

		$request = new Google_Service_Sheets_BatchUpdateSpreadsheetRequest( $request_body );
		$result  = $this->service->spreadsheets->batchUpdate( $spreadsheet_id, $request );

	}

	/**
	 * Create styles for columns
	 *
	 * @since 1.0.0
	 *
	 * @param $spreadsheet_id
	 * @param $items_count
	 * @param $customer_lines_count
	 */
	private function setFontStyles( $spreadsheet_id, $items_count, $customer_lines_count ) {

		$requests = [
			//Title
			new Google_Service_Sheets_Request( [
				'repeatCell' => [
					"range" => [
						"sheetId"          => 0,
						"startRowIndex"    => 0,
						"endRowIndex"      => 2,
						"startColumnIndex" => 1,
						"endColumnIndex"   => 3,
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
			//Print date
			new Google_Service_Sheets_Request( [
				'repeatCell' => [

					"range" => [
						"sheetId"          => 0,
						"startRowIndex"    => 0,
						"endRowIndex"      => 4,
						"startColumnIndex" => 3,
						"endColumnIndex"   => 4
					],
					"cell"  => [
						"userEnteredFormat" => [
							"horizontalAlignment" => "RIGHT",
							"textFormat"          => [
								"bold"     => true,
								"fontSize" => 10,
							]
						]
					],

					"fields" => "UserEnteredFormat(horizontalAlignment,textFormat)"
				]
			] ),
			//Print date values
			new Google_Service_Sheets_Request( [
				'repeatCell' => [

					"range" => [
						"sheetId"          => 0,
						"startRowIndex"    => 0,
						"endRowIndex"      => 4,
						"startColumnIndex" => 4,
						"endColumnIndex"   => 5
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
			//Font size and position for Customer Details
			new Google_Service_Sheets_Request( [
				'repeatCell' => [

					"range" => [
						"sheetId"          => 0,
						"startRowIndex"    => 2,
						"endRowIndex"      => $customer_lines_count,
						"startColumnIndex" => 1,
						"endColumnIndex"   => 2
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
						"startColumnIndex" => 2,
						"endColumnIndex"   => 3
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
			//Space before Specifications title
			new Google_Service_Sheets_Request( [
				'repeatCell' => [

					"range" => [
						"sheetId"          => 0,
						"startRowIndex"    => $customer_lines_count,
						"endRowIndex"      => $customer_lines_count + 1,
						"startColumnIndex" => 1,
						"endColumnIndex"   => 2,
					],
					"cell"  => [
						"userEnteredFormat" => [
							"horizontalAlignment" => "LEFT",
							"textFormat"          => [
								"bold"     => false,
								"fontSize" => 32,
							]
						]
					],

					"fields" => "UserEnteredFormat(horizontalAlignment,textFormat)"
				]
			] ),
			//Specifications title
			new Google_Service_Sheets_Request( [
				'repeatCell' => [

					"range" => [
						"sheetId"          => 0,
						"startRowIndex"    => $customer_lines_count + 1,
						"endRowIndex"      => $customer_lines_count + 2,
						"startColumnIndex" => 1,
						"endColumnIndex"   => 2,
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
			//Specifications attribute values
			new Google_Service_Sheets_Request( [
				'repeatCell' => [

					"range" => [
						"sheetId"          => 0,
						"startRowIndex"    => $customer_lines_count + 2,
						"endRowIndex"      => $items_count + $customer_lines_count + 2,
						"startColumnIndex" => 2,
						"endColumnIndex"   => 3,
					],
					"cell"  => [
						"userEnteredFormat" => [
							"horizontalAlignment" => "LEFT",
							"textFormat"          => [
								"bold" => false,
							]
						]
					],

					"fields" => "UserEnteredFormat(horizontalAlignment,textFormat)"
				]
			] ),
			//Main border
			new Google_Service_Sheets_Request( [
				'updateBorders' => [
					"range" => [
						"sheetId"          => 0,
						"startRowIndex"    => $customer_lines_count + 2,
						"endRowIndex"      => $items_count + $customer_lines_count + 2,
						"startColumnIndex" => 1,
						"endColumnIndex"   => 5
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
			//Bottom border
			new Google_Service_Sheets_Request( [
				'updateBorders' => [
					"range" => [
						"sheetId"          => 0,
						"startRowIndex"    => $items_count + $customer_lines_count + 2,
						"endRowIndex"      => $items_count + $customer_lines_count + 4,
						"startColumnIndex" => 1,
						"endColumnIndex"   => 5
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
				'method_id',
				'cost',
			)
		);
	}

	/**
	 * Check if sheet exist
	 *
	 * @since 1.0.0
	 *
	 * @param $spreadsheet_id
	 *
	 * @return bool
	 * @throws \Google\Exception
	 */
	public function check_sheet( $spreadsheet_id ) {

		$client  = $this->getGoogleClient();
		$service = new Google_Service_Sheets( $client );

		try {
			$response = $service->spreadsheets->get( $spreadsheet_id );

			if ( ! $response || empty( $response ) ) {
				return false;
			}

		} catch ( Exception $e ) {
			return false;
		}

		return true;

	}

	/**
	 * Make ordering for multiply array
	 *
	 * @param array $meta_data
	 *
	 * @return array
	 */
	private function prepare_reorganized_attributes( array $meta_data ) {

		$attributes                             = array();
		$attributes['_uni_cpo_threshold_type']  = array();
		$attributes['_uni_cpo_threshold_color'] = array();
		$attributes['_uni_cpo_handle_prep']     = array();
		$attributes['_uni_cpo_swing_config']    = array();
		$attributes['_uni_cpo_glass_color_92']  = array();
		$attributes['_uni_cpo_pivot_placement'] = array();
		$attributes['_uni_cpo_door_thickness']  = array();

		foreach ( $meta_data as $meta ) {
			if ( isset( $attributes[ $meta->key ] ) ) {
				$attributes[ $meta->key ] = $meta;
			}
		}

		return $attributes;
	}

}