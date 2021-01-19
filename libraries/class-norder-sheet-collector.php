<?php

/**
 * Prepare data from order and send it to sheets
 *
 * @since 1.0.0
 */
defined( 'ABSPATH' ) || exit;

class NOrderSheet_Collector {

	/**
	 * NOrderSheet_Collector constructor.
	 *
	 * @since 1.0.0
	 */
	public function __construct() {

		add_action( 'admin_enqueue_scripts', array( $this, 'add_scripts' ) );
		add_action( 'manage_shop_order_posts_custom_column', array( $this, 'add_google_sheet_button' ) );
		add_action( 'init', array( $this, 'catch_action' ) );

		add_filter( 'manage_edit-shop_order_columns', array( $this, 'add_google_sheet_column' ) );

	}

	/**
	 * Load scripts
	 *
	 * @since  1.0.0
	 */
	public function add_scripts() {

		wp_enqueue_script(
			'n-order-sheet-scripts',
			NORDER_SHEET_PLUGIN_URL . 'assets/scripts.js',
			array( 'jquery', 'jquery-get-xpath' ), NORDER_SHEET_VERSION,
			true
		);

		wp_enqueue_style(
			'n-order-sheet-styles',
			NORDER_SHEET_PLUGIN_URL . 'assets/styles.css',
			array(),
			NORDER_SHEET_VERSION
		);

	}

	/**
	 * Add column in orders for place export buttons
	 *
	 * @since 1.0.0
	 *
	 * @param $columns
	 *
	 * @return mixed
	 */
	public function add_google_sheet_column( $columns ) {

		$columns['n_order_sheet'] = __( 'Google Sheet', 'n-order-sheet' );

		return $columns;
	}

	/**
	 * Add button in new column
	 *
	 * @since 1.0.0
	 *
	 * @param $column
	 */
	public function add_google_sheet_button( $column ) {
		global $post;

		if ( 'n_order_sheet' === $column ) {

			$sheet_created = get_post_meta( $post->ID, 'norder_sheet_url', true );
			if ( $sheet_created ) { ?>
                <a href="<?php
				echo esc_url( add_query_arg( array(
					'norder_create_sheet_action' => 'view',
					'order_id'                   => $post->ID
				) ) ); ?>" class="order-status status-processing"
                   target="_blank">
                    <span><?php _e( 'View', 'n-order-sheet' ); ?></span>
                </a>
			<?php } else { ?>
                <a href="<?php
				echo esc_url( add_query_arg( array(
					'norder_create_sheet_action' => 'create',
					'order_id'                   => $post->ID
				) ) ); ?>" class="order-status status-processing" target="_blank">
                    <span><?php _e( 'Create', 'n-order-sheet' ); ?></span>
                </a>

			<?php }
		}
	}

	/**
	 * Listen actions and if needed, run it
	 *
	 * @since 1.0.0
	 */
	public function catch_action() {

		if ( ! isset( $_GET['norder_create_sheet_action'] ) || ! isset( $_GET['order_id'] ) ) {
			return false;
		}

		if ( 'create' == $_GET['norder_create_sheet_action'] ) {
			$this->build_order_sheet( $_GET['order_id'] );
		} else if ( 'view' == $_GET['norder_create_sheet_action'] ) {
			$this->check_sheet( $_GET['order_id'] );
		}

	}

	/**
	 * Build new sheet from order
	 *
	 * @since 1.0.0
	 *
	 * @param $order_id
	 *
	 * @throws Exception
	 */
	private function build_order_sheet( $order_id ) {

		$sheet_api      = new NOrderSheet_API();
		$spreadsheet_id = $sheet_api->create_sheet( $order_id );
		if ( $spreadsheet_id ) {
			$sheet_api->put_order_data( $spreadsheet_id, $order_id );
			$sheet_link = 'https://docs.google.com/spreadsheets/d/' . esc_attr( $spreadsheet_id ) . '/edit#gid=0';
			update_post_meta( $order_id, 'norder_sheet_url', $spreadsheet_id );
			?>
            <script language="JavaScript">
                document.location.href = '<?php echo esc_url( $sheet_link ); ?>';
            </script>
			<?php
			wp_die();
		}

	}

	/**
	 * Check if sheet is exist, we can redirect to him
	 *
	 * @since 1.0.0
	 *
	 * @param $order_id
	 *
	 * @throws Exception
	 */
	private function check_sheet( $order_id ) {

		$spreadsheet_id = get_post_meta( $order_id, 'norder_sheet_url', true );

		if ( $spreadsheet_id ) {
			$sheet_api = new NOrderSheet_API();

			if ( $sheet_api->check_sheet( $spreadsheet_id ) ) {
				$sheet_link = 'https://docs.google.com/spreadsheets/d/' . esc_attr( $spreadsheet_id ) . '/edit#gid=0';
				?>
                <script language="JavaScript">
                    document.location.href = '<?php echo esc_url( $sheet_link ); ?>';
                </script>
				<?php
				wp_die();
			}
		}

		$this->build_order_sheet( $order_id );

	}

	/**
	 * View early build sheet
	 *
	 * @since 1.0.0
	 *
	 * @param $order_id
	 */
	private function redirect_to_sheet( $order_id ) {

	}

}

function norder_sheet_collector_runner() {

	return new NOrderSheet_Collector();
}

norder_sheet_collector_runner();