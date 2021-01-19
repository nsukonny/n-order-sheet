<?php
/**
 * Plugin Name: N-Order Sheet
 * Description: WooCommerce Extension send order data to google sheets
 * Version: 1.0.0
 * Text Domain: n-order-sheet
 * Domain Path: /languages
 * WC requires at least: 3.0.0
 * WC tested up to: 4.4.1
 */

defined( 'ABSPATH' ) || exit;

if ( ! class_exists( 'NOrderSheet' ) ) {

	include_once dirname( __FILE__ ) . '/libraries/class-norder-sheet.php';

}

/**
 * The main function for returning NOrderSheet instance
 *
 * @since 1.0.0
 *
 * @return object The one and only true NOrderSheet instance.
 */
function norder_sheet_runner() {

	return NOrderSheet::instance();
}

norder_sheet_runner();

include_once dirname( __FILE__ ) . '/libraries/class-norder-sheet-activation.php';
register_activation_hook( __FILE__, array( 'NOrderSheet_Activation', 'activation' ) );
register_deactivation_hook( __FILE__, array( 'NOrderSheet_Activation', 'deactivate' ) );
