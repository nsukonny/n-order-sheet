<?php

/**
 * Settings page for this plugin
 *
 * @since 1.0.0
 */
defined( 'ABSPATH' ) || exit;

class NOrderSheet_Settings {

	/**
	 * NOrderSheet_Settings constructor.
	 *
	 * @since 1.0.0
	 */
	public function __construct() {

		add_action( 'admin_menu', array( $this, 'add_link_to_menu' ), 10, 1 );
		add_action( 'plugin_action_links_' . plugin_basename( __FILE__ ), array( $this, 'add_link_in_plugins_list' ) );

		if ( isset( $_GET['act'] ) && 'remove_token' == $_GET['act'] ) {
			NOrderSheet_API::deleteAccessToken();
			add_action( 'admin_notices', array( $this, 'token_success_deleted' ) );
		}

		if ( isset( $_GET['act'] ) && 'clear_sheets' == $_GET['act'] ) {
			delete_post_meta_by_key( 'norder_sheet_url' );
			add_action( 'admin_notices', array( $this, 'sheets_success_unlinked' ) );
		}

	}

	/**
	 * Make link to settings page
	 *
	 * @since 1.0.0
	 */
	public function add_link_to_menu() {

		add_menu_page(
			__( 'N-Order Sheet', 'n-order-sheet' ),
			__( 'N-Order Sheet', 'n-order-sheet' ),
			'manage_options',
			'n-order-sheet',
			array(
				$this,
				'display_settings',
			),
			'dashicons-tickets-alt',
			75
		);

	}

	/**
	 * Render settings
	 *
	 * @since 1.0.0
	 *
	 * @return void
	 * @throws Exception
	 */
	public function display_settings() {

		?>
        <div class="n-order-settings">
            <h1>N Order Sheet Settings</h1>
			<?php
			$api  = new NOrderSheet_API();
			$form = $api->getClient();

			?>
        </div>
		<?php

	}

	/**
	 * Add plugin action links.
	 *
	 * Add a link to the settings page on the plugins.php page.
	 *
	 * @since 1.0.0
	 *
	 * @param  array $links List of existing plugin action links.
	 *
	 * @return array         List of modified plugin action links.
	 */
	public function add_link_in_plugins_list( $links ) {

		$links = array_merge( array(
			'<a href="' . esc_url( admin_url( '/admin.php?page=n-order-sheet' ) ) . '">' . __( 'Settings', 'n-order-sheet' ) . '</a>'
		), $links );

		return $links;

	}

	/**
	 * Show message about deleting token
	 *
	 * @since 1.0.0
	 */
	public function token_success_deleted() {
		?>
        <div class="notice notice-success is-dismissible">
            <p><?php _e( 'Token successful deleted', 'n-order-sheet' ); ?></p>
        </div>
		<?php
	}

	/**
	 * Unlink all sheets from orders
	 *
	 * @since 1.0.0
	 */
	public function sheets_success_unlinked() {
		?>
        <div class="notice notice-success is-dismissible">
            <p><?php _e( 'All sheets unlinked from orders but still isset in your Google Drive', 'n-order-sheet' ); ?></p>
        </div>
		<?php
	}

}

function norder_sheet_settings_runner() {

	return new NOrderSheet_Settings();
}

norder_sheet_settings_runner();