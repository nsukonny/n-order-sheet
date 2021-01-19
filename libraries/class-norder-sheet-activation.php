<?php

/**
 * Static methods for installation and uninstall hooks
 *
 * @since 1.0.0
 */
defined( 'ABSPATH' ) || exit;

class NOrderSheet_Activation {

	/**
	 * Make some actions when admin will activating the plugin
	 *
	 * @since 1.0.0
	 */
	public static function activation() {

		$plugin_main_file = plugin_dir_path( dirname( __FILE__ ) ) . 'n-order-sheet.php';

		register_uninstall_hook( $plugin_main_file, array( 'NOrderSheet_Activation', 'uninstall' ) );

	}

	/**
	 * Clear all data when we uninstall the plugin
	 *
	 * @since 1.0.0
	 */
	public static function uninstall() {

		delete_post_meta_by_key( 'norder_sheet_url' );

	}

	/**
	 * Deactivation plugin
	 *
	 * @since 1.0.0
	 */
	public static function deactivate() {

		delete_post_meta_by_key( 'norder_sheet_url' );
		if ( file_exists( NORDER_SHEET_TOKEN_PATH ) ) {
			unlink( NORDER_SHEET_TOKEN_PATH );
		}

	}
}