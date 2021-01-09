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
	function add_link_in_plugins_list( $links ) {

		$links = array_merge( array(
			'<a href="' . esc_url( admin_url( '/admin.php?page=n-order-sheet' ) ) . '">' . __( 'Settings', 'n-order-sheet' ) . '</a>'
		), $links );

		return $links;

	}

}

function norder_sheet_settings_runner() {

	return new NOrderSheet_Settings();
}

norder_sheet_settings_runner();