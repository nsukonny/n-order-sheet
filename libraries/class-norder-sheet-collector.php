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

			$order         = wc_get_order( $post->ID );
			$sheet_created = get_post_meta( $post->ID, 'sheet_created', true );
			if ( $sheet_created ) { ?>
                <mark class="order-status status-processing">
                    <span>
                        <a href="<?php echo esc_url( add_query_arg( array(
	                        'norder_create_sheet_action' => 'view',
	                        'order_id'                   => $post->ID
                        ) ) ); ?>">
                            <?php _e( 'View', 'n-order-sheet' ); ?>
                        </a>
                    </span>
                </mark>
			<?php } else { ?>
                <mark class="order-status status-processing">
                    <span>
                        <a href="<?php echo esc_url( add_query_arg( array(
	                        'norder_create_sheet_action' => 'create',
	                        'order_id'                   => $post->ID
                        ) ) ); ?>">
                            <?php _e( 'Create', 'n-order-sheet' ); ?>
                        </a>
                    </span>
                </mark>
			<?php }
		}
	}

	/**
	 * Listen actions and if needed, run it
	 *
	 * @since 1.0.0
	 */
	public function catch_action() {

		if ( ! isset( $_GET['norder_create_sheet_action'] ) || ! isset( $_GET['orderId'] ) ) {
			return false;
		}

		if ( 'create' == $_GET['norder_create_sheet_action'] ) {
			$this->build_order_sheet( $_GET['orderId'] );
		} else if ( 'view' == $_GET['norder_create_sheet_action'] ) {
			$this->redirect_to_sheet( $_GET['orderId'] );
		}

	}

	/**
	 * Build new sheet from order
	 *
	 * @since 1.0.0
	 *
	 * @param $order_id
	 */
	private function build_order_sheet( $order_id ) {

		$order = wc_get_order( $order_id );

		echo '<h2>Shop Copy Order Specifications Sheet</h2>';
		echo '<br>Customer - ' . $order->get_formatted_billing_full_name();
		echo '<br>Phone - ' . $order->get_billing_phone();
		echo '<br>Shipping contact - ' . $order->get_formatted_shipping_full_name();
		echo '<br>Shipping phone - ' . $order->get_meta( 'shipping_phone', true );
		echo '<br>Shipping phone - ' . $order->get_formatted_shipping_address();

		echo '<br>Invoice number - ????';
		echo '<br>Scheduled ship date - ????';

		echo '<br><br><h3>Specifications</h3>';

		$hidden_order_itemmeta = apply_filters(
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

		foreach ( $order->get_items() as $item_id => $item ) {

			$product_id   = $item->get_product_id();
			$variation_id = $item->get_variation_id();
			$product      = $item->get_product();
			$name         = $item->get_name();
			$quantity     = $item->get_quantity();
			$subtotal     = $item->get_subtotal();
			$total        = $item->get_total();
			$tax          = $item->get_subtotal_tax();
			$taxclass     = $item->get_tax_class();
			$taxstat      = $item->get_tax_status();
			$allmeta      = $item->get_meta_data();
			$type         = $item->get_type();


			$somemeta = $item->get_meta( 'finish-option', true );

			echo '<br>' . $item->get_quantity() . ' ' . $product->get_title();
			if ( $somemeta ) {
				echo ' - ' . $somemeta;
			}

			if ( $meta_data = $item->get_formatted_meta_data( '' ) ) : ?>
                <table cellspacing="0" class="display_meta">
					<?php
					foreach ( $meta_data as $meta_id => $meta ) :
						if ( in_array( $meta->key, $hidden_order_itemmeta, true ) ) {
							continue;
						}
						?>
                        <tr>
                            <th><?php echo wp_kses_post( $meta->display_key ); ?>:</th>
                            <td><?php echo wp_kses_post( force_balance_tags( $meta->display_value ) ); ?></td>
                        </tr>
					<?php endforeach; ?>
                </table>
			<?php endif;
		}

		wp_die();
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