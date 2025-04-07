#!/bin/bash
#
# updated 5/3/2019 - DD & DL
#
# Script to triage a ChromeOS or Chromium OS drive and identify possible evidence of interest for an examiner.
#
#


# If we're not running as root, restart as root.
if [ ${UID:-$(id -u)} -ne 0 ]; then
  exec sudo "$0" "$@"
fi

# Load functions and constants for chromeos-install.
. /usr/share/misc/chromeos-common.sh || exit 1
. /usr/sbin/write_gpt.sh || exit 1

# Like mount but keeps track of the current mounts so that they can be cleaned
# up automatically.
tracked_mount() {
  local last_arg
  eval last_arg=\$$#
  MOUNTS="${last_arg}${MOUNTS:+ }${MOUNTS:-}"
  mount "$@"
}

# Unmount with tracking.
tracked_umount() {
  # dash does not support ${//} expansions.
  local new_mounts
  for mount in $MOUNTS; do
    if [ "$mount" != "$1" ]; then
      new_mounts="${new_mounts:-}${new_mounts+ }$mount"
    fi
  done
  MOUNTS=${new_mounts:-}

  umount "$1"
  rmdir "$1"
}

# Create a loop device on the given file at a specified (sector) offset.
# Remember the loop device using the global variable LOOP_DEV.
# Invoke as: command
# Args: FILE OFFSET BLKSIZE
loop_offset_setup() {
  local filename=$1
  local offset=$2
  local blocksize=$3

  if [ "${blocksize}" -eq 512 ]; then
    local param=""
  else
    local param="-b ${blocksize}"
  fi

  LOOP_DEV=$(losetup -f ${param} --show -o $(($offset * blocksize)) ${filename})
  if [ -z "$LOOP_DEV" ]; then
    die "No free loop device. Free up a loop device or reboot. Exiting."
  fi

  LOOPS="${LOOP_DEV}${LOOPS:+ }${LOOPS:-}"
}

# Delete the current loop device.
loop_offset_cleanup() {
  # dash does not support ${//} expansions.
  local new_loops
  for loop in $LOOPS; do
    if [ "$loop" != "$LOOP_DEV" ]; then
      new_loops="${new_loops:-}${new_loops+ }$loop"
    fi
  done
  LOOPS=${new_loops:-}

  # losetup -a doesn't always show every active device, so we'll always try to
  # delete what we think is the active one without checking first. Report
  # success no matter what.
  losetup -d ${LOOP_DEV} || /bin/true
}

# Mount the existing loop device at the mountpoint in $TMPMNT.
# Args: optional 'readwrite'. If present, mount read-write, otherwise read-only.
mount_on_loop_dev() {
  local rw_flag=${1-readonly}
  local mount_flags=""
  # if [ "${rw_flag}" != "readwrite" ]; then
     mount_flags="-o ro"
  # fi
  tracked_mount ${mount_flags} ${LOOP_DEV} ${TMPMNT}
}

# Unmount loop-mounted device.
umount_from_loop_dev() {
  mount | grep -q " on ${TMPMNT} " && tracked_umount ${TMPMNT}
}

# Undo all mounts and loops.
cleanup() {
  set +e

  local mount_point
  for mount_point in ${MOUNTS:-}; do
    umount "$mount_point" || /bin/true
  done
  MOUNTS=""

  local loop_dev
  for loop_dev in ${LOOPS:-}; do
    losetup -d "$loop_dev" || /bin/true
  done
  LOOPS=""
}


main() {
 
# Set variable for the root (USB boot) device. 
ROOTDEV=$(rootdev -d)

#
# Time for the user to select which drive they wish to triage!
#

echo "Please make sure you have disconnected any USB drive other than your boot drive or a previously created evidence clone of a ChromeOS or Chromium OS device."
echo "Hit Enter to continue when you have disconnected any additional USB drives!" 
read continue 

readarray -t lines < <(lsblk --nodeps -no name,vendor,model,serial,size,type | grep disk) 

# Prompt the user to select a drive.
echo "Please select the evidence drive that contains the stateful (STATE) GPT partition you wish to triage."
echo " "
select choice in "${lines[@]}"; do
[[ -n $choice ]] || { echo "Invalid choice. Please try again." >&2; continue; }
break # valid choice was made; exit prompt.
done

# Split out the ID of the selected drive.
read -r id unused <<<"$choice"

echo ${choice}
DEV=/dev/${id}
echo ${DEV}

if [[ $DEV == $ROOTDEV ]]; then
	echo "You have selected the OS boot drive you are booted from right now (i.e. Your currently running ChromeOS or Chromium OS), not an evidence drive."
	echo "Please re-run this script and select an evidence drive...exiting script, without doing anything!"
	echo ""	
	return 1
else

	# Set a variable for the internal HD device containing the STATE partition, such as /dev/mmcblk0p1.
	 DST=$(cgpt find -l STATE ${DEV})
	 # echo -n "DST:"
	 # echo ${DST}

	# Set variable for block device of DST. This removes the partition identifier from the above command, leaving /dev/sdb.
	 BLOCK_DST=$(get_block_dev_from_partition_dev ${DST})
	 # echo -n "Destination Block Device: "
	 # echo ${BLOCK_DST} 
	 
	# Find the partition number of the STATE partition on the internal HD and set a variable for it, such as 1.
	 PARTITION_NUM_STATE=$(cgpt find -n -l STATE "${DEV}")
	 # echo -n "PARTITION_NUM_STATE:"
	 # echo ${PARTITION_NUM_STATE}
	 
	# Create a temp folder to be used for a mount point later and set a variable for the mount point.
	 TMPMNT=$(mktemp -d)

	# Set variable for base device name from the STATE partition of internal HD, to be fed into blocksize function to determine block size, such as /dev/mmcblk0p1 (-> mmcblk0p1.
	 BASE_DST=$(basename ${DST})
	 # echo -n "BASE_DST:"
	 # echo ${BASE_DST}
	 
	# Set variable for block size of the internal HD that contains the STATE partition, such as 512. 
	 DST_BLKSIZE=$(blocksize ${BASE_DST})
	 # echo -n "DST_BLKSIZE:"
	 # echo ${DST_BLKSIZE}
	 
	# Extract the whole disk block device from the partition device.
	# This works for /dev/sda3 -> /dev/sda -> sda as well as /dev/mmcblk0p2 -> /dev/mmcblk0 -> mmcblk0 and set it to a variable.
	 BLOCK=$(get_block_dev_from_partition_dev ${DST##*/})
	 # echo -n "BLOCK:"
	 # echo ${BLOCK}

	# Set variable for starting offset of STATE partition of internal HD.
	 STATE_OFFSET=$(cgpt show -b -i ${PARTITION_NUM_STATE} ${BLOCK_DST})
	 # echo -n "STATE_OFFSET:"
	 # echo ${STATE_OFFSET}

	echo ""
	echo "Mounting (readonly) the STATE partition of the selected evidence drive..."
	loop_offset_setup ${BLOCK_DST} ${STATE_OFFSET} ${DST_BLKSIZE}
	mount_on_loop_dev readonly

    #
	# Identify all encrypted user vault folders on the selected drive.
	#
	echo "" 	
	echo "The encrypted user vaults that are present on " ${DEV} " are..."
	echo ""	
	# find ${TMPMNT}/home/.shadow/ -name "????????????????????????????????????????"
	
	x=0
	vault_list=()
	while IFS= read folder ; do
	    vault_list=("${vault_list[@]}" "$folder")
		done < <(find ${TMPMNT}/home/.shadow/ -maxdepth 1 -name "????????????????????????????????????????" -exec basename {} \;)
		echo "Vault list:"
		printf '%s\n' "${vault_list[@]}"
		((x++))

	#
	# Collect any known Google Account usernames!
	#

	#Prompt user for any Google Account they know existed on this Chromebook/Chrombox.
	while true
	do
	  # (1) prompt user, and read command line argument
	  echo "" 
	  echo "Do you know the username (i.e. email address) of any user accounts that were used on this Chromebook/Chromebox? (Y/N) "
	  read -n 1 ANSWER
	  echo ""
	  i=0
	  # (2) handle the input we were given
	  case $ANSWER in
	   [yY]* )  echo "Type each of the known username(s) separated by a space and hit ENTER."
				echo "Goggle does not use special characters in the prefix of email addresses"
				echo "So you must type first.last@gmail.com as firstlast@gmail.com!"	
				read -a CROSUSERNAME
				echo ""	
				c=0 
				for i in ${CROSUSERNAME[@]}
					do
					# cryptohome --action=dump_keyset --user=${CROSUSERNAME[$c]}	
					eval "CONCAT[$c]=$(echo -n $i | cat ${TMPMNT}/home/.shadow/salt -)"	
					OBFUSCATEUSER[$c]=$(echo -n ${CONCAT[$c]} | sha1sum | cut -d " " -f 1)
					echo "Username: " $i " obfuscated is " ${OBFUSCATEUSER[$c]}
						
						count=0
						while [ $count -le $x ]
						do
							((count++))
							if [ "${OBFUSCATEUSER[$c]}" = "${vault_list[$c]}" ]; then
								echo ""
								echo "Username " $i " owns the encrypted vault " ${vault_list[$c]}
								break
							fi
						done
						
					((c++))
					done	 
				break;;

	   [nN]* ) break;;

	   * )     echo "Please just enter Y or N.";;
	  esac

	done

	#
	# Now enumerate the files in each user encrypted vault.
	#
	
	count=0
	for j in ${vault_list[@]}
		do
		echo ""
		echo "The encrypted vault " ${vault_list[$count]} " contains the following files and folders:"
		echo ""
		
		# Might want to verify these are size 1
		USERFLDR[$count]=$(find ${TMPMNT}/home/.shadow/${vault_list[$count]}/mount -mindepth 1 -maxdepth 1 -gid 1001)
		CUTSIZE=$(tr -dc '/' <<< ${USERFLDR[$count]} | wc -c)
		find ${USERFLDR[$count]} -mindepth 1 -perm 710 -type d | cut -d/ -f${CUTSIZE} | uniq -u
		# FOLDERLIST[$count]=$(find ${USERFLDR[$count]} -mindepth 1 -perm 710 -type d | cut -d/ -f$(expr $CUTSIZE + 2) | uniq -u)
		DOWNLOADSFLDR[$count]=$(find ${USERFLDR[$count]} -mindepth 1 -perm 710 -type d | cut -d/ -f$(expr $CUTSIZE + 2) | uniq -u)
		echo ""
		echo "This is the encrypted Downloads folder: "	
		printf '%s\n' "${DOWNLOADSFLDR[$count]}"
		# DOWNLOADSFLDR[$count]=$(printf '%s\n' "${FOLDERLIST[$count]}")
		echo ""
		if [ ${#DOWNLOADSFLDR[$count]} -le 21 ]; then
			DOWNLOADSCONTENTS[$count]=$(ls -al ${USERFLDR[$count]}/*)
			echo "This is either a decrypted/mounted ChromeOS disk or a Chromium OS disk, so we are listing all"
			echo "subfolders and files in the user folder for the user vault " ${vault_list[$count]}
			echo "The vault contains the following: "	
		else	
			DOWNLOADSCONTENTS[$count]=$(ls -al ${USERFLDR[$count]}/${DOWNLOADSFLDR[$count]}/*)
			echo "The encrypted Downloads folder of the user vault " ${vault_list[$count]} " contains the following: "
		fi
		echo ""
		printf '%s\n' "${DOWNLOADSCONTENTS[$count]}"	
		((count++))
	done




	umount_from_loop_dev 
	sync
	loop_offset_cleanup
fi
 
 
 # All done.
 sync
 cleanup
 trap - EXIT

 echo "------------------------------------------------------------"
 echo ""
 echo "Triage complete."
 echo "If you didn't run this command with . triage_stateful.sh | tee output.txt,  to save all output to a text logfile"
 echo " then you may want to run it again, using the tee command to output all activity for a record of the results!"


}

main "$@"




