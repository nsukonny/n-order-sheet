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

		register_uninstall_hook( __FILE__, array( 'NOrderSheet_Activation', 'uninstall' ) );

	}

	/**
	 * Clear all data when we uninstall the plugin
	 *
	 * @since 1.0.0
	 */
	public static function uninstall() {

		delete_post_meta_by_key( 'norder_sheet_url' );

	}
}